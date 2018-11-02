use std::ffi::OsStr;
use std::mem;

use comical::com::{create_instance_local_server, getter};
use comical::error::{LabelErrorHResult, Result, check_hresult};
use winapi::shared::guiddef::GUID;
use winapi::shared::minwindef::ULONG;
use winapi::um::bits::{BackgroundCopyManager, BG_ERROR_CONTEXT, BG_JOB_PROGRESS, BG_JOB_STATE, BG_JOB_STATE_ERROR, BG_JOB_STATE_TRANSIENT_ERROR, BG_JOB_TYPE_DOWNLOAD, IBackgroundCopyJob, IBackgroundCopyManager};
use winapi::um::winnt::HRESULT;
use wio::com::ComPtr;
use wio::wide::ToWide;

pub fn connect_bcm() -> Result<ComPtr<IBackgroundCopyManager>> {
    create_instance_local_server::<BackgroundCopyManager, IBackgroundCopyManager>()
        .map_api_hr("CoCreateInstance")
}

pub fn create_download_job(display_name: &OsStr) -> Result<(GUID, ComPtr<IBackgroundCopyJob>)>
{
    let bcm = connect_bcm()?;
    let mut guid = unsafe { mem::uninitialized() };
    let job = getter(|job| unsafe {
        bcm.CreateJob(
            display_name.to_wide_null().as_ptr(),
            BG_JOB_TYPE_DOWNLOAD,
            &mut guid,
            job,
            )}).map_api_hr("IBackgroundCopyManager::CreateJob")?;
    Ok((guid, job))
}

pub fn get_job(guid: &GUID) -> Result<ComPtr<IBackgroundCopyJob>> {
    let bcm = connect_bcm()?;
    getter(|job| unsafe { bcm.GetJob(guid, job) }).map_api_hr("IBackgroundCopyManager::GetJob")
}

pub trait BitsJob {
    fn add_file(&self, remote_url: &OsStr, local_file: &OsStr) -> Result<()>;
    fn resume(&self) -> Result<()>;
    fn suspend(&self) -> Result<()>;
    fn complete(&self) -> Result<()>;
    fn cancel(&self) -> Result<()>;
    fn get_status(&self) -> Result<BitsJobStatus>;
}

pub struct BitsJobError {
    pub context: BG_ERROR_CONTEXT,
    pub error: HRESULT,
}

pub struct BitsJobStatus {
    pub state: BG_JOB_STATE,
    pub progress: BG_JOB_PROGRESS,
    pub error_count: ULONG,
    pub error: Option<BitsJobError>,
}

impl BitsJob for ComPtr<IBackgroundCopyJob> {
    fn add_file(&self, remote_url: &OsStr, local_file: &OsStr) -> Result<()>
    {
        check_hresult(unsafe {
            self.AddFile(
                remote_url.to_wide_null().as_ptr(),
                local_file.to_wide_null().as_ptr(),
            )}).map_api_hr("IBackgroundCopyJob::AddFile")?;
        Ok(())
    }

    fn resume(&self) -> Result<()> {
        check_hresult(unsafe { self.Resume() }).map_api_hr("IBackgroundCopyJob::Resume")?;
        Ok(())
    }

    fn suspend(&self) -> Result<()> {
        check_hresult(unsafe { self.Suspend() }).map_api_hr("IBackgroundCopyJob::Suspend")?;
        Ok(())
    }

    fn complete(&self) -> Result<()> {
        check_hresult(unsafe { self.Complete() }).map_api_hr("IBackgroundCopyJob::Complete")?;
        Ok(())
    }

    fn cancel(&self) -> Result<()> {
        check_hresult(unsafe { self.Cancel() }).map_api_hr("IBackgroundCopyJob::Cancel")?;
        Ok(())
    }

    fn get_status(&self) -> Result<BitsJobStatus> {
        let mut state = 0;
        check_hresult(unsafe { self.GetState(&mut state) })
            .map_api_hr("IBackgroundCopyJob::GetState")?;

        let mut progress = unsafe { mem::uninitialized() };
        check_hresult(unsafe { self.GetProgress(&mut progress) })
            .map_api_hr("IBackgroundCopyJob::GetProcess")?;

        let mut error_count = 0;
        check_hresult(unsafe { self.GetErrorCount(&mut error_count) })
            .map_api_hr("IBackgroundCopyJob::GetErrorCount")?;

        Ok(BitsJobStatus {
            state,
            progress,
            error_count,
            error: if state == BG_JOB_STATE_ERROR || state == BG_JOB_STATE_TRANSIENT_ERROR {
                let error_obj = getter(|e| unsafe { self.GetError(e) })
                    .map_api_hr("IBackgroundCopyJob::GetError")?;

                let mut context = 0;
                let mut hresult = 0;
                check_hresult(unsafe { error_obj.GetError(&mut context, &mut hresult) })
                    .map_api_hr("IBackgroundCopyError::GetError")?;

                Some(BitsJobError {
                    context,
                    error: hresult,
                })
            } else {
                None
            }
        })
    }
}
