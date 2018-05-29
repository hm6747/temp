function upLoadFile(ele,fileName) {
    var file =ele.files[0];
    var formData = new FormData();
    formData.append("file" , file);
    formData.append("fileName",fileName+".xlsx");
    $.ajax({
        url: '/upload',
        type: 'POST',
        data: formData,
        cache: false,
        contentType: false,
        processData: false,
        success: function (result) {
            if(result.code == 0){
                $(ele).attr("name",result.data)
            }else{
                alert(result.msg);
            }
        },
        error: function (result) {
            alert(result.msg);
        }
    });
}

function submitExcel(type) {
    var comprehensive = $("#comprehensive").attr("name");
    var signUp = $("#signUp").attr("name");
    var purchasingDemand = $("#purchasingDemand").attr("name");
    var PreliminaryReviewReport = $("#PreliminaryReviewReport").attr("name");
    var data = {
        comprehensive:comprehensive,
        signUp:signUp,
        purchasingDemand:purchasingDemand,
        PreliminaryReviewReport:PreliminaryReviewReport,
        type:type
    }
    $.ajax({
        type: "post",
        url: "/excel",
        data: data,
        dataType: "json",
        success: function (result) {
            if(result.code==0){

            }else{
                alert(result.msg);
            }
        },
        error: function (result) {
            alert(result.msg);
        }

    })
}