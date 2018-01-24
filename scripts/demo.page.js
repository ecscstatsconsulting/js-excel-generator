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
});
