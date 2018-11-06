use std::ffi::OsStr;
use std::mem;

use comical::com::{create_instance_local_server, getter};
use comical::error::{check_hresult, LabelErrorHResult, Result};
use comical::guid::Guid;
use winapi::shared::minwindef::ULONG;
use winapi::um::bits::{
    BackgroundCopyManager, IBackgroundCopyError, IBackgroundCopyJob, IBackgroundCopyManager,
    BG_ERROR_CONTEXT, BG_JOB_PRIORITY, BG_JOB_PROGRESS, BG_JOB_STATE, BG_JOB_STATE_ERROR,
    BG_JOB_STATE_TRANSIENT_ERROR, BG_JOB_TYPE_DOWNLOAD,
};
use winapi::um::winnt::HRESULT;
use wio::com::ComPtr;
use wio::wide::ToWide;

use comical::{call, get};

pub fn connect_bcm() -> Result<ComPtr<IBackgroundCopyManager>> {
    create_instance_local_server::<BackgroundCopyManager, IBackgroundCopyManager>()
}

pub fn create_download_job(display_name: &OsStr) -> Result<(Guid, ComPtr<IBackgroundCopyJob>)> {
    let bcm = connect_bcm()?;
    let mut guid = unsafe { mem::uninitialized() };
    let job = unsafe {
        get!(
            |job| bcm,
            IBackgroundCopyManager::CreateJob(
                display_name.to_wide_null().as_ptr(),
                BG_JOB_TYPE_DOWNLOAD,
                &mut guid,
                job,
            )
        )
    }?;

    Ok((Guid(guid), job))
}

pub fn get_job(guid: &Guid) -> Result<ComPtr<IBackgroundCopyJob>> {
    let bcm = connect_bcm()?;
    unsafe { get!(|job| bcm, IBackgroundCopyManager::GetJob(&guid.0, job)) }
}

pub trait BitsJob {
    fn add_file(&self, remote_url: &OsStr, local_file: &OsStr) -> Result<()>;
    fn set_description(&self, description: &OsStr) -> Result<()>;
    fn set_priority(&self, priority: BG_JOB_PRIORITY) -> Result<()>;
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
    fn add_file(&self, remote_url: &OsStr, local_file: &OsStr) -> Result<()> {
        unsafe {
            call!(
                self,
                IBackgroundCopyJob::AddFile(
                    remote_url.to_wide_null().as_ptr(),
                    local_file.to_wide_null().as_ptr(),
                )
            )
        }?;
        Ok(())
    }

    fn set_description(&self, description: &OsStr) -> Result<()> {
        unsafe {
            call!(
                self,
                IBackgroundCopyJob::SetDescription(description.to_wide_null().as_ptr())
            )
        }?;
        Ok(())
    }

    // TODO
    //fn set_proxy()

    fn set_priority(&self, priority: BG_JOB_PRIORITY) -> Result<()> {
        unsafe { call!(self, IBackgroundCopyJob::SetPriority(priority)) }?;
        Ok(())
    }

    fn resume(&self) -> Result<()> {
        unsafe { call!(self, IBackgroundCopyJob::Resume()) }?;
        Ok(())
    }

    fn suspend(&self) -> Result<()> {
        unsafe { call!(self, IBackgroundCopyJob::Suspend()) }?;
        Ok(())
    }

    fn complete(&self) -> Result<()> {
        unsafe { call!(self, IBackgroundCopyJob::Complete()) }?;
        Ok(())
    }

    fn cancel(&self) -> Result<()> {
        unsafe { call!(self, IBackgroundCopyJob::Cancel()) }?;
        Ok(())
    }

    fn get_status(&self) -> Result<BitsJobStatus> {
        let mut state = 0;
        let mut progress = unsafe { mem::uninitialized() };
        let mut error_count = 0;

        unsafe {
            call!(self, IBackgroundCopyJob::GetState(&mut state))?;
            call!(self, IBackgroundCopyJob::GetProgress(&mut progress))?;
            call!(self, IBackgroundCopyJob::GetErrorCount(&mut error_count))?;
        }

        Ok(BitsJobStatus {
            state,
            progress,
            error_count,
            error: if state == BG_JOB_STATE_ERROR || state == BG_JOB_STATE_TRANSIENT_ERROR {
                let error_obj = unsafe { get!(|e| self, IBackgroundCopyJob::GetError(e)) }?;

                let mut context = 0;
                let mut hresult = 0;
                unsafe {
                    call!(
                        error_obj,
                        IBackgroundCopyError::GetError(&mut context, &mut hresult)
                    )
                }?;

                Some(BitsJobError {
                    context,
                    error: hresult,
                })
            } else {
                None
            },
        })
    }
}
