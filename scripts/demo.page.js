$(document).ready(function () {
    $("#generate-excel-basic").click(function () {
		excel = new ExcelGen({
			"src_id": "basic_table",
			"show_header": true,
			"type": "table"
		});
        excel.generate();
    });
    $("#generate-excel-formatted").click(function () {
		excel = new ExcelGen({
			"src_id": "formatted_table",
			"show_header": true,
			"type": "table",
			"column_formats": ["1", "1", "4", "6", "18"] 
		});
        excel.generate();
	});
    $("#generate-csv").click(function () {
		excel = new ExcelGen({
			"file_name": "output.csv",
			"src_id": "csv_table",
			"show_header": true,
			"format": "csv"
		});
        excel.generate();
	});
	var myTable = $("#exclude_table");
	$("#generate-exclude").click(function () {
		excel = new ExcelGen({
			"file_name": "output.xlsx",
			"src": myTable,
			"show_header": true,
			"exclude_selector": ".xl_none"
		});
		excel.generate();
	});
});
