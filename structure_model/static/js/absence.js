$(document).ready(function() {
    $('#absenceForm').submit(function(e) {
        e.preventDefault();

        const formData = {
            tab_no: $('input[name="tab_no"]').val(),
            shift: $('select[name="shift"]').val(),
            day: $('input[name="day"]').val(),
            reason: $('select[name="reason"]').val()
        };

        $.ajax({
            url: '/submit-absence',
            method: 'POST',
            contentType: 'application/json',
            data: JSON.stringify(formData),
            success: function(response) {
                if (response.success) {
                    $('#formResponse').html(
                        '<div class="alert alert-success">' + response.message + '</div>'
                    );

                    $('#absenceForm')[0].reset();

                    const currentMode = $('#modeSwitch').val();
                    updateAbsenceStats(currentMode);
                    loadRecentAbsences();

                    setTimeout(function() {
                        const modalElement = document.getElementById('absenceModal');
                        const modalInstance = bootstrap.Modal.getInstance(modalElement);
                        if (modalInstance) {
                            modalInstance.hide();
                        }
                    }, 2000);
                }
            },
            error: function(xhr) {
                const error = (xhr.responseJSON && xhr.responseJSON.error) ? xhr.responseJSON.error : 'Ошибка сервера';
                $('#formResponse').html(
                    '<div class="alert alert-danger">' + error + '</div>'
                );
            }
        });
    });
});