use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use std::ptr::null_mut;

use comical::bstr::{bstr_from_u16, BStr};
use comical::com::{cast, check_nonzero, getter};
use comical::create_instance;
use comical::handle::check_hresult;
use comical::safearray::SafeArray;
use comical::variant::{Variant, VariantValue, VARIANT_FALSE, VARIANT_TRUE};

use winapi::shared::minwindef::{DWORD, MAX_PATH};
use winapi::shared::ntdef::LONG;
use winapi::shared::winerror::{ERROR_FILE_NOT_FOUND, HRESULT_FROM_WIN32};
use winapi::um::processthreadsapi::GetCurrentProcess;
use winapi::um::taskschd::{
    IExecAction, IRegisteredTask, ITaskFolder, ITaskService, TaskScheduler, TASK_ACTION_EXEC,
    TASK_CREATE_OR_UPDATE, TASK_DONT_ADD_PRINCIPAL_ACE, TASK_INSTANCES_IGNORE_NEW,
    TASK_LOGON_SERVICE_ACCOUNT,
};
use winapi::um::winbase::QueryFullProcessImageNameW;
use wio::com::ComPtr;

fn connect_task_service() -> Result<(ComPtr<ITaskService>, ComPtr<ITaskFolder>), String> {
    let task_service = create_instance!(TaskScheduler, ITaskService)?;
    check_hresult("ITaskService::Connect", unsafe {
        let null = Variant::null().get();
        task_service.Connect(
            null, // serverName
            null, // user
            null, // domain
            null, // password
        )
    })?;
    let root_folder = getter("Get Task Scheduler root folder", |root_folder| unsafe {
        task_service.GetFolder(BStr::from("\\").get(), root_folder)
    })?;

    Ok((task_service, root_folder))
}

fn get_task(task_path: &BStr) -> Result<Option<ComPtr<IRegisteredTask>>, String> {
    let root_folder = connect_task_service()?.1;
    let mut task = null_mut();

    let hr = unsafe { root_folder.GetTask(task_path.get(), &mut task as *mut *mut _) };
    if hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) {
        Ok(None)
    } else {
        check_hresult("GetTask", hr)?;
        Ok(Some(unsafe { ComPtr::from_raw(task) }))
    }
}

pub fn install(task_name: &str) -> Result<(), String> {
    let task_name = BStr::from(task_name);
    let mut image_path = [0u16; MAX_PATH + 1];
    check_nonzero("QueryFullProcessImageNameW", unsafe {
        let mut image_path_size = (image_path.len() - 1) as DWORD;
        QueryFullProcessImageNameW(
            GetCurrentProcess(),
            0, // dwFlags
            image_path.as_mut_ptr(),
            &mut image_path_size as *mut _,
        )
    })?;

    let task_def;
    let root_folder;
    {
        let (task_service, rf) = connect_task_service()?;
        root_folder = rf;

        // If the same task exists, remove it. Allowed to fail.
        unsafe { root_folder.DeleteTask(task_name.get(), 0) };

        task_def = getter("Create new task", |task_def| unsafe {
            task_service.NewTask(
                0, // flags (reserved)
                task_def,
            )
        })?;
    }

    {
        let reg_info = getter("get_RegistrationInfo", |info| unsafe {
            task_def.get_RegistrationInfo(info)
        })?;

        check_hresult("put_Author", unsafe {
            reg_info.put_Author(BStr::from("Mozilla").get())
        })?;
    }

    {
        let settings = getter("get_Settings", |s| unsafe { task_def.get_Settings(s) })?;

        check_hresult("put_MultipleInstances", unsafe {
            settings.put_MultipleInstances(TASK_INSTANCES_IGNORE_NEW)
        })?;

        check_hresult("put_AllowDemandStart", unsafe {
            settings.put_AllowDemandStart(VARIANT_TRUE)
        })?;

        check_hresult("put_RunOnlyIfIdle", unsafe {
            settings.put_RunOnlyIfIdle(VARIANT_FALSE)
        })?;

        check_hresult("put_DisallowStartIfOnBatteries", unsafe {
            settings.put_DisallowStartIfOnBatteries(VARIANT_FALSE)
        })?;

        check_hresult("put_StopIfGoingOnBatteries", unsafe {
            settings.put_StopIfGoingOnBatteries(VARIANT_FALSE)
        })?;

        let idle_settings = getter("get_IdleSettings", |s| unsafe {
            settings.get_IdleSettings(s)
        })?;

        check_hresult("put_StopOnIdleEnd", unsafe {
            idle_settings.put_StopOnIdleEnd(VARIANT_FALSE)
        })?;
    }

    {
        let action_collection = getter("get_Actions", |ac| unsafe { task_def.get_Actions(ac) })?;

        let action = getter("Create Action", |a| unsafe {
            action_collection.Create(TASK_ACTION_EXEC, a)
        })?;

        let exec_action = cast::<_, IExecAction>(action, "IExecAction")?;

        check_hresult("Set exec action path", unsafe {
            exec_action.put_Path(bstr_from_u16(&image_path).get())
        })?;

        check_hresult("Set exec action args", unsafe {
            exec_action.put_Arguments(BStr::from("task $(Arg0) $(Arg1)").get())
        })?;
    }

    let registered_task = getter("RegisterTaskDefinition", |rt| unsafe {
        root_folder.RegisterTaskDefinition(
            task_name.get(),
            task_def.as_raw(),
            TASK_CREATE_OR_UPDATE as LONG,
            Variant::<BStr>::wrap(&mut BStr::from("NT AUTHORITY\\LocalService")).get(),
            Variant::null().get(), // password
            TASK_LOGON_SERVICE_ACCOUNT,
            Variant::<BStr>::wrap(&mut BStr::empty()).get(), // sddl
            rt,
        )
    })?;

    // Allow read and execute access by builtin users, this is required to Get the task and
    // call Run on it
    // TODO: should this just be in sddl above? I think that ends up adding BU as principal?
    check_hresult("SetSecurityDescriptor", unsafe {
        registered_task.SetSecurityDescriptor(
            BStr::from("D:(A;;GRGX;;;BU)").get(),
            TASK_DONT_ADD_PRINCIPAL_ACE as LONG,
        )
    })?;

    Ok(())
}

pub fn uninstall(task_name: &str) -> Result<(), String> {
    let task_name = BStr::from(task_name);
    let (_, root_folder) = connect_task_service()?;
    check_hresult("DeleteTask", unsafe {
        root_folder.DeleteTask(task_name.get(), 0)
    })?;

    Ok(())
}

pub fn run_on_demand(task_name: &str, args: &[OsString]) -> Result<(), String> {
    let task_name = BStr::from(task_name);

    let args = args
        .iter()
        .map(|a| bstr_from_u16(&a.encode_wide().collect::<Vec<u16>>()))
        .collect::<Vec<_>>();

    let maybe_task = get_task(&task_name)?;
    if maybe_task.is_none() {
        return Err(String::from("No such task"));
    }

    let mut sa = SafeArray::try_from(args)?;
    let v = Variant::<SafeArray<_>>::wrap(&mut sa)?;
    if let VariantValue::StringVector(Ok(v)) = v.value() {
        for (i, s) in v.iter().enumerate() {
            println!("{}: {}", i, s);
        }
    }
    getter("Run Task", |rt| unsafe {
        maybe_task.unwrap().Run(v.get(), rt)
    })?;

    Ok(())
}
