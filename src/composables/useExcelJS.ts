import { ref } from "vue";
import ExcelJS, { Cell, Row, Workbook, Worksheet } from "exceljs";
import { IKey, IDirection } from "../types";

export const useExcelJS = () => {
  const workbook = ref<Workbook | null>(null);

  const importExcelFile = async (event: Event) => {
    const target = event.target as HTMLInputElement;
    const file = target.files ? target.files[0] : null;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (e: ProgressEvent<FileReader>) => {
      if (e.target?.result) {
        const data = new Uint8Array(e.target.result as ArrayBuffer);
        workbook.value = new ExcelJS.Workbook();
        await workbook.value.xlsx.load(data);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportExcelFile = async (data: any[], direction: IDirection) => {
    if (!workbook.value) {
      alert("No file loaded");
      return;
    }

    const newWorkbook = new ExcelJS.Workbook();
    const keys = ref<IKey[]>([]);

    // Var olan sayfaları yeni Workbook'a kopyala
    workbook.value.worksheets.forEach((sheet) => {
      const newSheet = newWorkbook.addWorksheet(sheet.name);

      sheet.eachRow({ includeEmpty: true }, (row: Row) => {
        const newRow = newSheet.addRow(row.values);

        row.eachCell({ includeEmpty: true }, (cell: Cell, colNumber) => {
          const newCell = newRow.getCell(colNumber) as any;
          newCell.value = cell.value ?? "";
          newCell.style = cell.style;

          data.forEach((item) => {
            Object.keys(item).forEach((key) => {
              if (
                typeof newCell.value === "string" &&
                newCell.value.includes(key) &&
                !keys.value.find((k) => k.name === key)
              ) {
                keys.value.push({
                  name: key,
                  col: newCell._column.number,
                  row: newCell._row.number,
                });
              }
            });
          });

          keys.value.forEach((key) => {
            if (direction === "vertical") {
              if (
                newCell._row.number - key.row >= 0 &&
                newCell._row.number - key.row < data.length &&
                newCell._column.number === key.col
              ) {
                newCell.value = data[newCell._row.number - key.row][key.name];
              }
            } else if (direction === "horizontal") {
              if (
                newCell._column.number - key.col >= 0 &&
                newCell._column.number - key.col < data.length &&
                newCell._row.number === key.row
              ) {
                newCell.value =
                  data[newCell._column.number - key.col][key.name];
              }
            }
          });
        });
      });
    });


    // Bir önceki sayfayı kopyala ve yeni bir sayfa olarak ekle
    // const lastSheet =
    //   workbook.value.worksheets[workbook.value.worksheets.length - 1]; // Son sayfayı seç
    // const copiedSheet = newWorkbook.addWorksheet(`${lastSheet.name} - Kopya`); // Yeni sayfa ekle
    // copySheet(lastSheet, copiedSheet);


    // Yeni Workbook'u dışa aktarma
    const buffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "exported-file.xlsx";
    link.click();
    window.URL.revokeObjectURL(url);
  };

  return {
    workbook,
    importExcelFile,
    exportExcelFile,
  };
};

// const copySheet = (sheet: Worksheet, copiedSheet: any) => {
//   // Son sayfadaki tüm satırları yeni sayfaya kopyala
//   sheet.eachRow({ includeEmpty: true }, (row: Row) => {
//     const newRow = copiedSheet.addRow(row.values);

//     row.eachCell({ includeEmpty: true }, (cell: Cell, colNumber) => {
//       const newCell = newRow.getCell(colNumber) as any;
//       newCell.value = cell.value ?? "";
//       newCell.style = cell.style;
//     });
//   });

//   // Sütun genişliklerini ve stillerini kopyala
//   if (sheet.columns) {
//     copiedSheet.columns = sheet.columns.map((col) => ({
//       width: col.width,
//       style: col.style,
//     }));
//   }

//   // Birleştirilmiş hücreleri kopyala
//   if (sheet.mergeCells && Array.isArray(sheet.mergeCells)) {
//     //@ts-ignore
//     for (const range of sheet.mergeCells) {
//       copiedSheet.mergeCells(range);
//     }
//   }
// };
