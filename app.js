Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run").onclick = run;
    }
});

async function run() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.getRange("A1").values = [["Hello, Excel!"]];
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}