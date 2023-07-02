function TDSCOllectionFileUploadSuccess(response) {
    debugger;
    if (response.status == true) {
        successAlert('File Has Been Succesfully Uploaded.');
    } else {
        errorAlert('File Has Not Been Uploaded. Please Try Again !!');
    }
}

function TDSCOllectionFileUploadComplete(response) {

}


function sendTDSCOllectionMailSuccess(resp) {
    debugger;
    if (resp.status == true) {
        successAlert('Mail Has been Sent Successfully.');
    } else {
        errorAlert('Mail Sending Failed. Please Try Again !!');
    }
}

function Loader(msg) {
    Swal.fire({
        title: 'Please Wait...',
        html: msg,
        allowOutsideClick: false,
        onBeforeOpen: () => {
            Swal.showLoading()
        },
    });
}

function successAlert(msg) {
    Swal.fire({
        title: 'Success!!',
        text: msg,
        type: 'success',
        confirmButtonColor: '#3085d6',
        confirmButtonText: 'OK!',
        allowOutsideClick: false
    });
}

function errorAlert(msg) {
    Swal.fire({
        title: 'Error!',
        text: msg,
        type: 'error',
        confirmButtonColor: '#3085d6',
        confirmButtonText: 'OK!',
        allowOutsideClick: false
    });
}