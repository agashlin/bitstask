use std::ffi::OsStr;
use std::mem;
use std::panic::RefUnwindSafe;

use comical::com::{create_instance_local_server, getter};
use comical::error::{check_hresult, LabelErrorHResult, Result};
use comical::guid::Guid;
use winapi::um::bits::{
    BackgroundCopyManager, IBackgroundCopyCallback, IBackgroundCopyError, IBackgroundCopyJob,
    IBackgroundCopyManager, BG_JOB_PRIORITY, BG_JOB_STATE_ERROR, BG_JOB_STATE_TRANSIENT_ERROR,
    BG_JOB_TYPE_DOWNLOAD, BG_NOTIFY_JOB_ERROR, BG_NOTIFY_JOB_MODIFICATION,
    BG_NOTIFY_JOB_TRANSFERRED,
};
use winapi::um::unknwnbase::IUnknown;
use wio::com::ComPtr;
use wio::wide::ToWide;

use comical::error::Error;
use comical::{call, get};

use protocol::{BitsJobError, BitsJobStatus};

pub type Callback = (Fn() -> () + RefUnwindSafe + Send + Sync + 'static);

pub fn connect_bcm() -> Result<ComPtr<IBackgroundCopyManager>> {
    create_instance_local_server::<BackgroundCopyManager, IBackgroundCopyManager>()
}

pub struct BitsJob {
    job: ComPtr<IBackgroundCopyJob>,
}

impl BitsJob {
    pub fn new(display_name: &OsStr) -> Result<Self> {
        let bcm = connect_bcm()?;
        unsafe {
            let mut guid = mem::uninitialized();
            let job = get!(
                |job| bcm,
                IBackgroundCopyManager::CreateJob(
                    display_name.to_wide_null().as_ptr(),
                    BG_JOB_TYPE_DOWNLOAD,
                    &mut guid,
                    job,
                )
            )?;

            Ok(BitsJob { job })
        }
    }

    pub fn get_by_guid(guid: &Guid) -> Result<BitsJob> {
        let bcm = connect_bcm()?;
        // TODO: something special for no such job error
        let job = unsafe { get!(|job| bcm, IBackgroundCopyManager::GetJob(&guid.0, job)) }?;

        Ok(BitsJob { job })
    }

    pub fn guid(&self) -> Result<Guid> {
        unsafe {
            let mut guid = mem::uninitialized();
            call!(self.job, IBackgroundCopyJob::GetId(&mut guid))?;
            Ok(Guid(guid))
        }
    }

    pub fn add_file(&mut self, remote_url: &OsStr, local_file: &OsStr) -> Result<()> {
        unsafe {
            call!(
                self.job,
                IBackgroundCopyJob::AddFile(
                    remote_url.to_wide_null().as_ptr(),
                    local_file.to_wide_null().as_ptr(),
                )
            )
        }?;
        Ok(())
    }

    pub fn set_description(&mut self, description: &OsStr) -> Result<()> {
        unsafe {
            call!(
                self.job,
                IBackgroundCopyJob::SetDescription(description.to_wide_null().as_ptr())
            )
        }?;
        Ok(())
    }

    // TODO
    //fn set_proxy()

    pub fn set_priority(&mut self, priority: BG_JOB_PRIORITY) -> Result<()> {
        unsafe { call!(self.job, IBackgroundCopyJob::SetPriority(priority)) }?;
        Ok(())
    }

    pub fn resume(&mut self) -> Result<()> {
        unsafe { call!(self.job, IBackgroundCopyJob::Resume()) }?;
        Ok(())
    }

    pub fn suspend(&mut self) -> Result<()> {
        unsafe { call!(self.job, IBackgroundCopyJob::Suspend()) }?;
        Ok(())
    }

    pub fn complete(&mut self) -> Result<()> {
        unsafe { call!(self.job, IBackgroundCopyJob::Complete()) }?;
        // TODO need to handle partial completion
        Ok(())
    }

    pub fn cancel(&mut self) -> Result<()> {
        unsafe { call!(self.job, IBackgroundCopyJob::Cancel()) }?;
        Ok(())
    }

    pub fn register_callbacks(
        &mut self,
        transferred: Option<Box<Callback>>,
        error: Option<Box<Callback>>,
        modification: Option<Box<Callback>>,
    ) -> Result<()>
where {
        // TODO check via GetNotifyInterface
        /*if self.callback.is_some() {
            return Err(Error::Message("callback already registered".to_string()));
        }*/

        unsafe {
            call!(
                self.job,
                IBackgroundCopyJob::SetNotifyFlags(
                    if transferred.is_some() {
                        BG_NOTIFY_JOB_TRANSFERRED
                    } else {
                        0
                    } | if error.is_some() {
                        BG_NOTIFY_JOB_ERROR
                    } else {
                        0
                    } | if modification.is_some() {
                        BG_NOTIFY_JOB_MODIFICATION
                    } else {
                        0
                    }
                )
            )?;
        }

        let callback = Box::new(callback::BackgroundCopyCallback {
            interface: IBackgroundCopyCallback {
                lpVtbl: &callback::VTBL,
            },
            transferred,
            error,
            modification,
        });

        // TODO: don't just leak, proper ref counting
        unsafe {
            call!(
                self.job,
                IBackgroundCopyJob::SetNotifyInterface(Box::leak(callback)
                    as *mut callback::BackgroundCopyCallback
                    as *mut IUnknown)
            )?;
        }
        Ok(())
    }

    pub fn get_status(&mut self) -> Result<BitsJobStatus> {
        let mut state = 0;
        let mut progress = unsafe { mem::uninitialized() };
        let mut error_count = 0;

        unsafe {
            call!(self.job, IBackgroundCopyJob::GetState(&mut state))?;
            call!(self.job, IBackgroundCopyJob::GetProgress(&mut progress))?;
            call!(
                self.job,
                IBackgroundCopyJob::GetErrorCount(&mut error_count)
            )?;
        }

        Ok(BitsJobStatus {
            state,
            progress,
            error_count,
            error: if state == BG_JOB_STATE_ERROR || state == BG_JOB_STATE_TRANSIENT_ERROR {
                let error_obj = unsafe { get!(|e| self.job, IBackgroundCopyJob::GetError(e)) }?;

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

mod callback {
    use std::borrow::BorrowMut;
    use std::panic::{catch_unwind, RefUnwindSafe};

    use comical::guid::Guid;
    use winapi::ctypes::c_void;
    use winapi::shared::guiddef::REFIID;
    use winapi::shared::minwindef::DWORD;
    use winapi::shared::ntdef::ULONG;
    use winapi::shared::winerror::{E_NOINTERFACE, HRESULT, NOERROR, S_OK};
    use winapi::um::bits::{
        IBackgroundCopyCallback, IBackgroundCopyCallbackVtbl, IBackgroundCopyError,
        IBackgroundCopyJob,
    };
    use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl};
    use winapi::Interface;

    use bits::Callback;

    #[repr(C)]
    pub struct BackgroundCopyCallback {
        pub interface: IBackgroundCopyCallback,
        pub transferred: Option<Box<Callback>>,
        pub error: Option<Box<Callback>>,
        pub modification: Option<Box<Callback>>,
    }

    extern "system" fn query_interface(
        This: *mut IUnknown,
        riid: REFIID,
        ppvObj: *mut *mut c_void,
    ) -> HRESULT {
        unsafe {
            if Guid(*riid) == Guid(IUnknown::uuidof())
                || Guid(*riid) == Guid(IBackgroundCopyCallback::uuidof())
            {
                addref(This);
                *ppvObj = This as *mut c_void;
                NOERROR
            } else {
                E_NOINTERFACE
            }
        }
    }

    extern "system" fn addref(_This: *mut IUnknown) -> ULONG {
        // TODO learn Rust synchronization
        1
    }

    extern "system" fn release(_This: *mut IUnknown) -> ULONG {
        // TODO
        1
    }

    extern "system" fn transferred_stub(
        This: *mut IBackgroundCopyCallback,
        _pJob: *mut IBackgroundCopyJob,
    ) -> HRESULT {
        unsafe {
            let this = This as *mut BackgroundCopyCallback;
            if let Some(cb) = (*this).transferred.as_ref() {
                catch_unwind(|| cb());
            }
        }
        S_OK
    }

    extern "system" fn error_stub(
        This: *mut IBackgroundCopyCallback,
        _pJob: *mut IBackgroundCopyJob,
        _pError: *mut IBackgroundCopyError,
    ) -> HRESULT {
        S_OK
    }

    extern "system" fn modification_stub(
        This: *mut IBackgroundCopyCallback,
        _pJob: *mut IBackgroundCopyJob,
        _dwReserved: DWORD,
    ) -> HRESULT {
        S_OK
    }

    pub static VTBL: IBackgroundCopyCallbackVtbl = IBackgroundCopyCallbackVtbl {
        parent: IUnknownVtbl {
            QueryInterface: query_interface,
            AddRef: addref,
            Release: release,
        },
        JobTransferred: transferred_stub,
        JobError: error_stub,
        JobModification: modification_stub,
    };
}
