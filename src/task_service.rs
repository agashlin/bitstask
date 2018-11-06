use std::ffi::{OsStr, OsString};

use comical::bstr::BStr;
use comical::com::{cast, create_instance_inproc_server, getter};
use comical::error::{
    check_hresult, check_nonzero, Error, ErrorCode, LabelErrorDWord, LabelErrorHResult, Result,
};
use comical::safearray::SafeArray;
use comical::variant::{Variant, VariantValue, VARIANT_FALSE, VARIANT_TRUE};
use comical::{call, get};

use winapi::shared::minwindef::{DWORD, MAX_PATH};
use winapi::shared::ntdef::LONG;
use winapi::shared::winerror::{ERROR_FILE_NOT_FOUND, HRESULT_FROM_WIN32};
use winapi::um::processthreadsapi::GetCurrentProcess;
use winapi::um::taskschd::{
    IActionCollection, IExecAction, IIdleSettings, IRegisteredTask, IRegistrationInfo,
    ITaskDefinition, ITaskFolder, ITaskService, ITaskSettings, TaskScheduler, TASK_ACTION_EXEC,
    TASK_CREATE_OR_UPDATE, TASK_DONT_ADD_PRINCIPAL_ACE, TASK_INSTANCES_PARALLEL,
    TASK_LOGON_SERVICE_ACCOUNT,
};
use winapi::um::winbase::QueryFullProcessImageNameW;
use wio::com::ComPtr;

fn connect_task_service() -> Result<(ComPtr<ITaskService>, ComPtr<ITaskFolder>)> {
    let task_service = create_instance_inproc_server::<TaskScheduler, ITaskService>()?;

    // Connect to local service with no credentials.
    unsafe {
        let null = Variant::null().get();
        call!(task_service, ITaskService::Connect(null, null, null, null))?;
    }

    let root_folder = unsafe {
        get!(
            |folder| task_service,
            ITaskService::GetFolder(BStr::from("\\").get(), folder)
        )
    }?;

    Ok((task_service, root_folder))
}

fn get_task(task_path: &BStr) -> Result<Option<ComPtr<IRegisteredTask>>> {
    let (_, root_folder) = connect_task_service()?;

    let task = unsafe {
        get!(
            |task| root_folder,
            ITaskFolder::GetTask(task_path.get(), task)
        )
    };
    match task {
        Ok(task) => Ok(Some(task)),
        Err(Error::Api(_, ErrorCode::HResult(hr), _))
            if hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) =>
        {
            Ok(None)
        }
        Err(e) => Err(e),
    }
}

pub fn install(task_name: &OsStr) -> Result<()> {
    let task_name = BStr::from(task_name);
    let mut image_path = [0u16; MAX_PATH + 1];
    let mut image_path_size_chars = (image_path.len() - 1) as DWORD;
    check_nonzero(unsafe {
        QueryFullProcessImageNameW(
            GetCurrentProcess(),
            0, // dwFlags
            image_path.as_mut_ptr(),
            &mut image_path_size_chars as *mut _,
        )
    }).map_api_rc("QueryFullProcessImageNameW")?;
    let image_path = &image_path[..image_path_size_chars as usize];

    let task_def;
    let root_folder;
    {
        let (task_service, rf) = connect_task_service()?;
        root_folder = rf;

        // If the same task exists, remove it. Allowed to fail.
        unsafe { root_folder.DeleteTask(task_name.get(), 0) };

        task_def = getter(|task_def| unsafe {
            task_service.NewTask(
                0, // flags (reserved)
                task_def,
            )
        }).map_api_hr("ITaskService::NewTask")?;
    }

    unsafe {
        let reg_info = get!(|info| task_def, ITaskDefinition::get_RegistrationInfo(info))?;
        call!(
            reg_info,
            IRegistrationInfo::put_Author(BStr::from("Mozilla").get())
        )?;
    }

    unsafe {
        let settings = get!(|s| task_def, ITaskDefinition::get_Settings(s))?;
        comical::call!(
            settings,
            ITaskSettings::put_MultipleInstances(TASK_INSTANCES_PARALLEL)
        )?;
        call!(settings, ITaskSettings::put_AllowDemandStart(VARIANT_TRUE))?;
        call!(settings, ITaskSettings::put_RunOnlyIfIdle(VARIANT_FALSE))?;
        call!(
            settings,
            ITaskSettings::put_DisallowStartIfOnBatteries(VARIANT_FALSE)
        )?;
        call!(
            settings,
            ITaskSettings::put_StopIfGoingOnBatteries(VARIANT_FALSE)
        )?;

        let idle_settings = get!(|is| settings, ITaskSettings::get_IdleSettings(is))?;
        call!(
            idle_settings,
            IIdleSettings::put_StopOnIdleEnd(VARIANT_FALSE)
        )?;
    }

    unsafe {
        let action_collection = get!(|ac| task_def, ITaskDefinition::get_Actions(ac))?;
        let exec_action = cast(get!(
            |a| action_collection,
            IActionCollection::Create(TASK_ACTION_EXEC, a)
        )?)?;
        call!(
            exec_action,
            IExecAction::put_Path(BStr::from(image_path).get())
        )?;
        call!(
            exec_action,
            IExecAction::put_Arguments(BStr::from("task $(Arg0) $(Arg1)").get())
        )?;
    }

    let registered_task = unsafe {
        get!(
            |rt| root_folder,
            ITaskFolder::RegisterTaskDefinition(
                task_name.get(),
                task_def.as_raw(),
                TASK_CREATE_OR_UPDATE as LONG,
                Variant::<BStr>::wrap(&mut BStr::from("NT AUTHORITY\\LocalService")).get(),
                Variant::null().get(), // password
                TASK_LOGON_SERVICE_ACCOUNT,
                Variant::<BStr>::wrap(&mut BStr::empty()).get(), // sddl
                rt,
            )
        )
    }?;

    // Allow read and execute access by builtin users, this is required to Get the task and
    // call Run on it
    // TODO: should this just be in sddl above? I think that ends up adding BU as principal?
    unsafe {
        call!(
            registered_task,
            IRegisteredTask::SetSecurityDescriptor(
                BStr::from("D:(A;;GRGX;;;BU)").get(),
                TASK_DONT_ADD_PRINCIPAL_ACE as LONG,
            )
        )
    }?;

    Ok(())
}

pub fn uninstall(task_name: &OsStr) -> Result<()> {
    let task_name = BStr::from(task_name);
    let (_, root_folder) = connect_task_service()?;
    unsafe {
        call!(
            root_folder,
            ITaskFolder::DeleteTask(BStr::from(task_name).get(), 0)
        )?
    };

    Ok(())
}

pub fn run_on_demand(task_name: &OsStr, args: &[OsString]) -> Result<()> {
    let task_name = BStr::from(task_name);

    let args = args
        .iter()
        .map(|a| BStr::from(a.as_os_str()))
        .collect::<Vec<_>>();

    let maybe_task = get_task(&task_name)?;
    if maybe_task.is_none() {
        return Err(Error::Message("No such task".to_string()));
    }
    let task = maybe_task.unwrap();

    // DEBUG
    let mut sa = SafeArray::try_from(args)?;
    let v = Variant::<SafeArray<_>>::wrap(&mut sa);
    if let VariantValue::StringVector(Ok(v)) = v.value() {
        for (i, s) in v.iter().enumerate() {
            println!("{}: {}", i, s);
        }
    }

    unsafe { get!(|rt| task, IRegisteredTask::Run(v.get(), rt))? };

    Ok(())
}
