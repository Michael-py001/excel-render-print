import React, { useState } from "react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

const Index = () => {
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState(null);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetNames = workbook.SheetNames;
      const sheetsData = sheetNames.map((name) => {
        const sheet = workbook.Sheets[name];
        const merges = sheet["!merges"] || [];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
        return { name, data, merges, sheet };
      });
      setSheets(sheetsData);
      setActiveSheet(sheetsData[0]);
    };

    reader.readAsArrayBuffer(file);
  };

  const handlePrint = () => {
    window.print();
  };

  const getCellStyle = (cell) => {
    if (!cell || !cell.s) return {};
    const { bgColor } = cell.s;
    return {
      backgroundColor: bgColor ? `#${bgColor.rgb.slice(2)}` : "transparent",
    };
  };

  const renderMergedCells = (sheet) => {
    const mergedCells = {};
    sheet.merges.forEach((merge) => {
      for (let R = merge.s.r; R <= merge.e.r; ++R) {
        for (let C = merge.s.c; C <= merge.e.c; ++C) {
          if (R === merge.s.r && C === merge.s.c) {
            mergedCells[`${R}-${C}`] = {
              rowspan: merge.e.r - merge.s.r + 1,
              colspan: merge.e.c - merge.s.c + 1,
            };
          } else {
            mergedCells[`${R}-${C}`] = { hidden: true };
          }
        }
      }
    });
    return mergedCells;
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
            {sheets.map((sheet) => {
              const mergedCells = renderMergedCells(sheet);
              return (
                <TabsContent key={sheet.name} value={sheet.name}>
                  <div className="overflow-auto">
                    <table className="min-w-full bg-white">
                      <tbody>
                        {sheet.data.map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            {row.map((cell, cellIndex) => {
                              const mergeInfo = mergedCells[`${rowIndex}-${cellIndex}`];
                              if (mergeInfo?.hidden) return null;
                              return (
                                <td
                                  key={cellIndex}
                                  className="border px-4 py-2"
                                  rowSpan={mergeInfo?.rowspan || 1}
                                  colSpan={mergeInfo?.colspan || 1}
                                  style={getCellStyle(sheet.sheet[XLSX.utils.encode_cell({ r: rowIndex, c: cellIndex })])}
                                >
                                  {cell}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </TabsContent>
              );
            })}
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