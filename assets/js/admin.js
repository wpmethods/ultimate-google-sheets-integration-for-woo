jQuery(document).ready(function($) {
    // Generate Google Apps Script
    $("#ugsiw-generate-script").on("click", function(e) {
        e.preventDefault();
        
        var button = $(this);
        var originalText = button.text();
        
        button.text("Generating...").prop("disabled", true);
        
        // Get selected fields
        var selectedFields = [];
        $("input[name='ugsiw_gs_selected_fields[]']:checked").each(function() {
            selectedFields.push($(this).val());
        });
        
        $.ajax({
            url: ugsiw_gs.ajax_url,
            type: "POST",
            data: {
                action: "ugsiw_generate_google_script",
                nonce: ugsiw_gs.nonce,
                fields: selectedFields
            },
            success: function(response) {
                if (response.success) {
                    $("#ugsiw-generated-script").val(response.data.script);
                    $("#ugsiw-script-output").show();
                    $("html, body").animate({
                        scrollTop: $("#ugsiw-script-output").offset().top - 100
                    }, 500);
                } else {
                    alert("Error generating script. Please try again.");
                }
            },
            error: function() {
                alert("Error generating script. Please try again.");
            },
            complete: function() {
                button.text(originalText).prop("disabled", false);
            }
        });
    });

    // Copy to clipboard
    $("#ugsiw-copy-script").on("click", function() {
        var textarea = $("#ugsiw-generated-script")[0];
        textarea.select();
        document.execCommand("copy");
        
        $("#ugsiw-copy-status").show().fadeOut(2000);
    });

    // Handle required fields
    $("input[name='ugsiw_gs_selected_fields[]']").each(function() {
        if ($(this).is(":disabled")) {
            $(this).prop("checked", true);
        }
    });

    // Prevent unchecking required fields
    $("input[name='ugsiw_gs_selected_fields[]']").on("change", function() {
        if ($(this).is(":disabled") && !$(this).is(":checked")) {
            $(this).prop("checked", true);
        }
    });

    // Toggle fields visibility
    $(".ugsiw-field-category h3").on("click", function() {
        $(this).next(".ugsiw-fields-grid").slideToggle(300);
        $(this).find(".dashicons").toggleClass("dashicons-arrow-down dashicons-arrow-up");
    });

    // Search fields
    $("#ugsiw-field-search").on("keyup", function() {
        var search = $(this).val().toLowerCase();
        $(".ugsiw-field-item").each(function() {
            var label = $(this).find("label").text().toLowerCase();
            if (label.indexOf(search) > -1) {
                $(this).show();
            } else {
                $(this).hide();
            }
        });
    });

    // Add/remove Apps Script URL rows (moved from inline PHP)
    $(document).on('click', '#ugsiw-add-script-url', function(e) {
        e.preventDefault();
        var row = jQuery('<div/>', { 'class': 'ugsiw-script-url-row', 'style': 'display:flex;gap:8px;margin-bottom:8px;' });
        row.append('<input type="url" name="ugsiw_gs_script_urls[]" placeholder="https://script.google.com/macros/s/..." style="flex:1;padding:10px;border:1px solid #ddd;border-radius:4px;">');
        row.append('<button type="button" class="button ugsiw-remove-script-url">Remove</button>');
        jQuery('#ugsiw-script-urls-list').append(row);
    });

    $(document).on('click', '.ugsiw-remove-script-url', function(e) {
        e.preventDefault();
        if ($('#ugsiw-script-urls-list .ugsiw-script-url-row').length > 1) {
            $(this).closest('.ugsiw-script-url-row').remove();
        } else {
            $(this).siblings('input').val('');
        }
    });
});