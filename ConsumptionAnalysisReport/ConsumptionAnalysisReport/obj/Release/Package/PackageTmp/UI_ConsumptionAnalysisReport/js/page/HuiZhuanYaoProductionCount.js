﻿////////////////////////////增加新变量/////////////////
var MaxValueArray = [];
var MinValueArray = [];
var MaxColumnCount = 9;
//////////////////////////////////////////////////////
$(function () {
    setDateBoxMonth();
    LoadDataGrid({ "rows": [], "total": 0 });
    QueryReportFun();
});

var m_SelectTime = '';
function QueryReportFun() {
    m_SelectTime = $('#startDate').datebox('getValue');
    var win = $.messager.progress({
        title: '请稍后',
        msg: '数据载入中...'
    });
    $.ajax({
        type: "POST",
        url: "HuiZhuanYaoProductionCount.aspx/GetConsumptionAnalysisInfo",
        data: '{m_SelectTime: "' + m_SelectTime + '"}',
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (msg) {
            $.messager.progress('close');
            var myData = jQuery.parseJSON(msg.d);
            SetMinMaxValue(myData);                //设置最大最小值
            //$('#DataGrid_ReportTable').datagrid("loadData", myData);
            LoadDataGrid(myData);
        },
        beforeSend: function (XMLHttpRequest) {
            win;
        },
        error: function () {
            $.messager.progress('close');
        }
    });
}

function LoadDataGrid(myData) {
    $('#DataGrid_ReportTable').datagrid({
        data: myData,
        dataType: "json",
        rownumbers: true,
        singleSelect: true,
        toolbar: '#toolbar_ReportTable',
        columns: [[
           { width: '120', title: '公司名称', field: 'CompanyName', rowspan: 2, align: 'center' },
           { width: '80', title: '生产线', field: 'ProductionLine', rowspan: 2, align: 'center'},
           { width: '320', title: '生料产量', colspan: 3, align: 'center' },
           { width: '320', title: '煤粉产量', colspan: 3, align: 'center' },
           { width: '320', title: '熟料产量', colspan: 3, align: 'center' }],
           [{ width: '120', title: '盘库', field: 'A1', align: 'center', styler: cellStyler_crosspurple },
            { width: '120', title: 'DCS', field: 'A2', align: 'center', styler: cellStyler_crosspurple },
            { width: '80', title: '增减比例', field: 'A3', align: 'center', styler: function (value, row, index) { return SetBackgroundColor(2, value, row, index); } },
            { width: '120', title: '盘库', field: 'A4', align: 'center', styler: cellStyler_crosspurple },
            { width: '120', title: 'DCS', field: 'A5', align: 'center', styler: cellStyler_crosspurple },
            { width: '80', title: '增减比例', field: 'A6', align: 'center', styler: function (value, row, index) { return SetBackgroundColor(5, value, row, index); } },
            { width: '120', title: '盘库', field: 'A7', align: 'center', styler: cellStyler_crosspurple },
            { width: '120', title: 'DCS', field: 'A8', align: 'center', styler: cellStyler_crosspurple },
            { width: '80', title: '增减比例', field: 'A9', align: 'center', styler: function (value, row, index) { return SetBackgroundColor(8, value, row, index); } }
        ]],
        onLoadSuccess: function (data) {
            var rowmark = 1;
            for (var i = 1; i < data.rows.length; i++) {
                if (data.rows[i]['CompanyName'] == data.rows[i - 1]['CompanyName']) {
                    rowmark += 1;
                    $(this).datagrid('mergeCells', {
                        index: i + 1 - rowmark,
                        field: 'CompanyName',  //合并单元格的区域，就是clomun中的filed对应的列
                        rowspan: rowmark
                    });
                }
                else {
                    rowmark = 1;
                }

            }
        }
    });
}

////////////////////////////设置最大最小值//////////////////////
function SetMinMaxValue(myData) {
    if (myData != undefined && myData != null && myData["rows"] != undefined && myData["rows"] != null) {
        for (var i = 0; i < MaxColumnCount; i++) {            //重新设置
            MaxValueArray[i] = -1;
            MinValueArray[i] = -1;
        }
        for (var i = 0; i < myData["rows"].length; i++) {
            for (var j = 0; j < MaxColumnCount; j++) {
                var m_ValueTemp = myData["rows"][i][(j + 1).toString()];
                if (m_ValueTemp != "" && m_ValueTemp != null) {
                    var m_FloatValueTemp = Math.abs(parseFloat(m_ValueTemp.replace("%", "")));//因为比较偏差,因此取绝对值
                    if (MaxValueArray[j] < m_FloatValueTemp || MaxValueArray[j] == -1) {
                        MaxValueArray[j] = m_FloatValueTemp;
                    }
                    if (MinValueArray[j] > m_FloatValueTemp || MinValueArray[j] == -1) {
                        MinValueArray[j] = m_FloatValueTemp;
                    }
                }
            }
        }
    }
}
//////////////////////////增加配色方案/////////////////////////

function SetBackgroundColor(myColumnIndex, value, row, index) {
    if (value != undefined && value != null) {
        var m_ValueTemp = Math.abs(parseFloat(value.replace("%")));
        //////////////////////本次是计算偏差,因此绝对值越小越好//////////////////
        if (m_ValueTemp == MaxValueArray[myColumnIndex]) {
            return 'background-color:#ff0000;';
        }
        else if (m_ValueTemp == MinValueArray[myColumnIndex]) {
            return 'background-color:#00ff00;';
        }
        else {
            return cellStyler_crosspurple(value, row, index);
        }
    }
    else {
        return cellStyler_crosspurple(value, row, index);
    }
}


function cellStyler_crossorange(value, row, index) {
    if (index % 2 == 0) {
        return 'background-color:#eeeeee;';
    }
    else {
        return 'background-color:#ffcc88;';
    }
}
function cellStyler_crosspurple(value, row, index) {
    if (index % 2 == 0) {
        return 'background-color:#eeeeee;';
    }
    else {
        return 'background-color:#e0c0ff;'
    }
}
function cellStyler_crossgreen(value, row, index) {
    if (index % 2 == 0) {
        return 'background-color:#eeeeee;';
    }
    else {
        return 'background-color:#ddffdd;'
    }
}
function cellStyler_crossred(value, row, index) {
    if (index % 2 == 0) {
        return 'background-color:#eeeeee;';
    }
    else {
        return 'background-color:#ffdddd;'
    }
}
function cellStyler_lightgray(value, row, index) {

    return 'background-color:#eeeeee;';
}

function cellStyler_gray(value, row, index) {

    return 'background-color:#cdb0a0;';
}
function cellStyler_Lightyellow(value, row, index) {

    return 'background-color:#ffff88;';
}
function cellStyler_Lightblue(value, row, index) {

    return 'background-color:#8888ff;';
}
function cellStyler_Lightpurple(value, row, index) {

    return 'background-color:#ff88cc;';
}
function cellStyler_Lightorange(value, row, index) {

    return 'background-color:#ffcc88;';
}

//////////////////////////////

function SetDivPosization(myWidth) {
    //var m_ContentDivWidth = $('#TableContentDiv').width() + 30;
    if (myWidth > 1188) {           //当窗口比div大
        $('#TableContentDiv').css('margin-left', parseInt((myWidth - 1188) / 2));
    }
    else {
        $('#TableContentDiv').css('margin-left', 0);
    }
}

function setDateBoxMonth() {
    $('#startDate').datebox({
        //显示日趋选择对象后再触发弹出月份层的事件，初始化时没有生成月份层
        onShowPanel: function () {
            //触发click事件弹出月份层
            span.trigger('click');
            if (!tds)
                //延时触发获取月份对象，因为上面的事件触发和对象生成有时间间隔
                setTimeout(function () {
                    tds = p.find('div.calendar-menu-month-inner td');
                    tds.click(function (e) {
                        //禁止冒泡执行easyui给月份绑定的事件
                        e.stopPropagation();
                        //得到年份
                        var year = /\d{4}/.exec(span.html())[0],
                        //月份
                        //之前是这样的month = parseInt($(this).attr('abbr'), 10) + 1; 
                        month = parseInt($(this).attr('abbr'), 10);

                        //隐藏日期对象                     
                        $('#startDate').datebox('hidePanel')
                          //设置日期的值
                          .datebox('setValue', year + '-' + month);
                    });
                }, 0);
        },
        //配置parser，返回选择的日期
        parser: function (s) {
            if (!s) return new Date();
            var arr = s.split('-');
            return new Date(parseInt(arr[0], 10), parseInt(arr[1], 10) - 1, 1);
        },
        //配置formatter，只返回年月 之前是这样的d.getFullYear() + '-' +(d.getMonth()); 
        formatter: function (d) {
            var currentMonth = (d.getMonth() + 1);
            var currentMonthStr = currentMonth < 10 ? ('0' + currentMonth) : (currentMonth + '');
            return d.getFullYear() + '-' + currentMonthStr;
        }
    });

    //日期选择对象
    var p = $('#startDate').datebox('panel'),
    //日期选择对象中月份
    tds = false,
    //显示月份层的触发控件
    span = p.find('span.calendar-text');
    var curr_time = new Date();

    //设置当月
    //var curr_time = new Date();
    //$("#startDate").datebox("setValue", myformatter(curr_time));

    //设置前一个月
    var nowDate = new Date();
    var oneMonthDate = new Date(nowDate - 30 * 24 * 3600 * 1000);
    $("#startDate").datebox("setValue", myformatter(oneMonthDate));
};

//格式化日期
function myformatter(date) {
    //获取年份
    var y = date.getFullYear();
    //获取月份
    var m = date.getMonth() + 1;
    return y + '-' + m;
}

function ExportFileFun() {
    var m_FunctionName = "ExcelStream";
    var m_Parameter1 = GetDataGridTableHtml("DataGrid_ReportTable", "回转窑系统产量统计", m_SelectTime);
    var m_Parameter2 = "";

    var m_ReplaceAlllt = new RegExp("<", "g");
    var m_ReplaceAllgt = new RegExp(">", "g");
    m_Parameter1 = m_Parameter1.replace(m_ReplaceAlllt, "&lt;");
    m_Parameter1 = m_Parameter1.replace(m_ReplaceAllgt, "&gt;");

    var form = $("<form id = 'ExportFile'>");   //定义一个form表单
    form.attr('style', 'display:none');   //在form表单中添加查询参数
    form.attr('target', '');
    form.attr('method', 'post');
    form.attr('action', "HuiZhuanYaoProductionCount.aspx");

    var input_Method = $('<input>');
    input_Method.attr('type', 'hidden');
    input_Method.attr('name', 'myFunctionName');
    input_Method.attr('value', m_FunctionName);
    var input_Data1 = $('<input>');
    input_Data1.attr('type', 'hidden');
    input_Data1.attr('name', 'myParameter1');
    input_Data1.attr('value', m_Parameter1);
    var input_Data2 = $('<input>');
    input_Data2.attr('type', 'hidden');
    input_Data2.attr('name', 'myParameter2');
    input_Data2.attr('value', m_Parameter2);

    $('body').append(form);  //将表单放置在web中 
    form.append(input_Method);   //将查询参数控件提交到表单上
    form.append(input_Data1);   //将查询参数控件提交到表单上
    form.append(input_Data2);   //将查询参数控件提交到表单上
    form.submit();
    //释放生成的资源
    form.remove();
}