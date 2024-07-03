import React, { useState } from "react";
import ExcelJS from "exceljs";
import { Button } from "@/components/ui/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

const Index = () => {
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState(null);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file);
    const sheetsData = workbook.worksheets.map((sheet) => {
      const data = [];
      sheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
        const rowData = [];
        row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
          rowData.push({
            value: cell.value,
            style: {
              backgroundColor: cell.fill?.fgColor?.argb ? `#${cell.fill.fgColor.argb.slice(2)}` : "transparent",
              fontSize: cell.font?.size ? `${cell.font.size}pt` : "inherit",
              fontWeight: cell.font?.bold ? "bold" : "normal",
            },
          });
        });
        data.push(rowData);
      });
      return { name: sheet.name, data };
    });
    setSheets(sheetsData);
    setActiveSheet(sheetsData[0]);
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="container mx-auto p-4">
      <h1 className="text-3xl mb-4">Excel Sheet Viewer</h1>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      {sheets.length > 0 && (
        <>
          <Tabs defaultValue={sheets[0].name} onValueChange={(value) => setActiveSheet(sheets.find(sheet => sheet.name === value))}>
            <TabsList>
              {sheets.map((sheet) => (
                <TabsTrigger key={sheet.name} value={sheet.name}>
                  {sheet.name}
                </TabsTrigger>
              ))}
            </TabsList>
            {sheets.map((sheet) => (
              <TabsContent key={sheet.name} value={sheet.name}>
                <div className="overflow-auto">
                  <table className="min-w-full bg-white">
                    <tbody>
                      {sheet.data.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {row.map((cell, cellIndex) => (
                            <td
                              key={cellIndex}
                              className="border px-4 py-2"
                              style={cell.style}
                            >
                              {cell.value}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </TabsContent>
            ))}
          </Tabs>
          <Button onClick={handlePrint} className="mt-4">
            Print
          </Button>
        </>
      )}
    </div>
  );
};

export default Index;