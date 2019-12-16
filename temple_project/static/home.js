$(document).ready(function () {
    let old_phone = "";
    let old_address = "";
    let set_button
    $("#id_phone").keyup(function () {

        $.ajax({

            url: '{% url "validate_date" %}',
            data: {
                "find_value": $("#id_phone").val()
            },
            dataType: 'json',
            success: function (data) {

                $("#data_table tbody").html("");

                for (let i = 0; i < data.find_format.length; i++) {
                    let split_data = data.find_format[i].split("/")
                    $('#data_table').append("<tr><td id=" + split_data[0] + ">" + split_data[0] + "</td><td id=" + split_data[1] + ">" + split_data[1] + "</td><td><button class=\"btn btn-info fix\" value=\"" + split_data[1] + "/" + split_data[0] + "\">修改 </button>   <button class=\"btn btn-info del\" value=\"" + split_data[0] + "\">刪除 </button></td></tr>")
                }
            }
        })
    });


    $('body').on('click', '.fix', function (e) {

        get_value = $(this).attr("value").split("/")
        old_address = get_value[0];
        old_phone = get_value[1];
        set_button = this
        $("#phone").val(old_phone);
        $("#address").val(old_address);
        $("#login_inputbox").dialog("open");
        e.preventDefault();

    });

    $('body').on('click', '.del', function () {

        if (confirm("確定刪除此家庭資料嗎？")) {
            get_value = $(this).attr("value")
            $.ajax({
                url: '{% url "validate_del" %}',

                data: {
                    "phone": get_value
                },
                dataType: 'json',
                success: function (data) {
                    if (data.is_taken) {
                        alert(data.result);
                    } else {
                        alert(data.error_message)
                    }
                }
            })
            $(this).closest("tr").remove()
        } else {
            alert("已取消")
        }
    });


    var check_login = function () {
        let new_phone = $("#phone").val()
        let new_address = $("#address").val()
        $.ajax({
            url: '{% url "validate_username" %}',

            data: {
                "old_phone": old_phone,
                "old_address": old_address,
                "new_phone": new_phone,
                "new_address": new_address
            },
            dataType: 'json',
            success: function (data) {
                if (data.is_taken) {
                    alert(data.result);
                    document.getElementById(old_address).innerHTML = new_address;
                    document.getElementById(old_address).id = new_address;
                    document.getElementById(old_phone).innerHTML = new_phone;
                    document.getElementById(old_phone).id = new_phone;
                    set_button.value = new_address + "/" + new_phone
                } else {
                    alert(data.error_message)
                }
            }
        })
        $(this).dialog("close");

    }


    $("#login_inputbox").dialog({

        width: 400,
        autoOpen: false,
        modal: true,
        title: "修改系統",
        buttons: {
            "送出": check_login,
            "取消": function () {
                $(this).dialog("close");
            }
        }
    });


})