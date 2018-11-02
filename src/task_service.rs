use std::ffi::{OsStr, OsString};

use comical::bstr::BStr;
use comical::com::{cast, create_instance_inproc_server, getter};
use comical::error::{
    check_hresult, check_nonzero, Error, ErrorCode, LabelErrorDWord, LabelErrorHResult, Result,
};
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

fn connect_task_service() -> Result<(ComPtr<ITaskService>, ComPtr<ITaskFolder>)> {
    let task_service = create_instance_inproc_server::<TaskScheduler, ITaskService>()
        .map_api_hr("CoCreateInstance")?;

    check_hresult(unsafe {
        let null = Variant::null().get();
        task_service.Connect(
            null, // serverName
            null, // user
            null, // domain
            null, // password
        )
    }).map_api_hr("Connect")?;
    let root_folder = getter(|root_folder| unsafe {
        task_service.GetFolder(BStr::from("\\").get(), root_folder)
    }).map_api_hr("GetFolder")?;

    Ok((task_service, root_folder))
}

fn get_task(task_path: &BStr) -> Result<Option<ComPtr<IRegisteredTask>>> {
    let (_, root_folder) = connect_task_service()?;

    match getter(|task| unsafe { root_folder.GetTask(task_path.get(), task) })
        .map_api_hr("GetTask")
    {
        Ok(task) => Ok(Some(task)),
        Err(Error::Api("GetTask", ErrorCode::HResult(hr)))
            if hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) =>
        {
            Ok(None)
        }
        Err(e) => Err(e),
    }
}

pub fn install(task_name: &OsStr) -> Result<()>
{
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
        }).map_api_hr("NewTask")?;
    }

    {
        let reg_info = getter(|info| unsafe { task_def.get_RegistrationInfo(info) })
            .map_api_hr("get_RegistrationInfo")?;

        check_hresult(unsafe { reg_info.put_Author(BStr::from("Mozilla").get()) })
            .map_api_hr("put_Author")?;
    }

    unsafe {
        let settings = getter(|s| task_def.get_Settings(s)).map_api_hr("get_Settings")?;
        check_hresult(settings.put_MultipleInstances(TASK_INSTANCES_IGNORE_NEW))
            .map_api_hr("put_MultipleInstances")?;
        check_hresult(settings.put_AllowDemandStart(VARIANT_TRUE))
            .map_api_hr("put_AllowDemandStart")?;
        check_hresult(settings.put_RunOnlyIfIdle(VARIANT_FALSE)).map_api_hr("put_RunOnlyIfIdle")?;
        check_hresult(settings.put_DisallowStartIfOnBatteries(VARIANT_FALSE))
            .map_api_hr("put_DisallowStartIfOnBatteries")?;
        check_hresult(settings.put_StopIfGoingOnBatteries(VARIANT_FALSE))
            .map_api_hr("put_StopIfGoingOnBatteries")?;

        let idle_settings =
            getter(|s| settings.get_IdleSettings(s)).map_api_hr("get_IdleSettings")?;
        check_hresult(idle_settings.put_StopOnIdleEnd(VARIANT_FALSE))
            .map_api_hr("put_StopOnIdleEnd")?;
    }

    unsafe {
        let action_collection = getter(|ac| task_def.get_Actions(ac)).map_api_hr("get_Actions")?;
        let action = getter(|a| action_collection.Create(TASK_ACTION_EXEC, a))
            .map_api_hr("IActionCollection::Create")?;
        let exec_action = cast::<_, IExecAction>(action)?;
        check_hresult(exec_action.put_Path(BStr::from(image_path).get()))
            .map_api_hr("put_Path")?;
        check_hresult(exec_action.put_Arguments(BStr::from("task $(Arg0) $(Arg1)").get()))
            .map_api_hr("put_Arguments")?;
    }

    let registered_task = getter(|rt| unsafe {
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
    }).map_api_hr("RegisterTaskDefinition")?;

    // Allow read and execute access by builtin users, this is required to Get the task and
    // call Run on it
    // TODO: should this just be in sddl above? I think that ends up adding BU as principal?
    check_hresult(unsafe {
        registered_task.SetSecurityDescriptor(
            BStr::from("D:(A;;GRGX;;;BU)").get(),
            TASK_DONT_ADD_PRINCIPAL_ACE as LONG,
        )
    }).map_api_hr("SetSecurityDescriptor")?;

    Ok(())
}

pub fn uninstall(task_name: &OsStr) -> Result<()> {
    let task_name = BStr::from(task_name);
    let (_, root_folder) = connect_task_service()?;
    check_hresult(unsafe { root_folder.DeleteTask(task_name.get(), 0) })
        .map_api_hr("DeleteTask")?;

    Ok(())
}

pub fn run_on_demand(task_name: &OsStr, args: &[OsString]) -> Result<()>
{
    let task_name = BStr::from(task_name);

    let args = args
        .iter()
        .map(|a| BStr::from(a.as_os_str()))
        .collect::<Vec<_>>();

    let maybe_task = get_task(&task_name)?;
    if maybe_task.is_none() {
        return Err(Error::Message("No such task".to_string()));
    }

    // DEBUG
    let mut sa = SafeArray::try_from(args)?;
    let v = Variant::<SafeArray<_>>::wrap(&mut sa);
    if let VariantValue::StringVector(Ok(v)) = v.value() {
        for (i, s) in v.iter().enumerate() {
            println!("{}: {}", i, s);
        }
    }

    getter(|rt| unsafe { maybe_task.unwrap().Run(v.get(), rt) }).map_api_hr("Run")?;

    Ok(())
}
