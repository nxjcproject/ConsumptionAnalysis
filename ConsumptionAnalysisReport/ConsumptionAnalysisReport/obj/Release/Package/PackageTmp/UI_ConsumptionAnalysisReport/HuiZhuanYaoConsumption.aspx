<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="HuiZhuanYaoConsumption.aspx.cs" Inherits="ConsumptionAnalysisReport.UI_ConsumptionAnalysisReport.HuiZhuanYaoConsumption" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>回转窑系统工序电耗对比</title>
    <link rel="stylesheet" type="text/css" href="/lib/ealib/themes/custom/easyui_datagrid_green.css" />
    <link rel="stylesheet" type="text/css" href="/lib/ealib/themes/icon.css" />
    <link rel="stylesheet" type="text/css" href="/lib/extlib/themes/syExtIcon.css" />

    <script type="text/javascript" src="/lib/ealib/jquery.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="/lib/ealib/jquery.easyui.min.js" charset="utf-8"></script>
    <script type="text/javascript" src="/lib/ealib/easyui-lang-zh_CN.js" charset="utf-8"></script>

    <!--[if lt IE 8 ]><script type="text/javascript" src="/js/common/json2.min.js"></script><![endif]-->
    <script type="text/javascript" src="/lib/ealib/extend/jquery.PrintArea.js" charset="utf-8"></script>
    <script type="text/javascript" src="/lib/ealib/extend/jquery.jqprint.js" charset="utf-8"></script>
    <script type="text/javascript" src="/js/common/PrintFile.js" charset="utf-8"></script> 

    <script type="text/javascript" src="js/page/HuiZhuanYaoConsumption.js" charset="utf-8"></script>
</head>
<body>
    <div id="Mainlayout" class="easyui-layout" data-options="fit:true,border:false">
        <div class="easyui-panel" data-options="region:'center',border:false,onResize:function(width,height){SetDivPosization(width);}" style="overflow: auto;">
            <div id="TableContentDiv">
                <div style="height: 28px;margin-top:2px;">
                    <span style="margin-left: 15px;">查询时间：<input id="startDate" type="text" class="easyui-datebox" required="required" style="width: 100px;" /></span>
                    <span style="margin-left: 10px;"><a href="javascript:void(0);" class="easyui-linkbutton" data-options="iconCls:'icon-search'" onclick="QueryReportFun();">查询</a></span>
                    <span style="margin-left: 10px;"><a href="#" class="easyui-linkbutton" data-options="iconCls:'ext-icon-page_white_excel',plain:true" onclick="ExportFileFun();">导出</a></span>
                </div>
                <div>
                    <div id="toolbar_ReportTable" style="display: none; text-align: center; vertical-align: middle; height: 30px; padding-top: 7px;">
                        <span style="font-size: 18pt; color:white; font-weight:bold; font-family:SimHei">回转窑系统工序电耗对比</span>
                    </div>
                    <div data-options="border:false">
                        <table id="DataGrid_ReportTable" style="width: 1303px;"></table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <form id="form1" runat="server"></form>
</body>
</html>
