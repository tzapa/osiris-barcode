import React, { useMemo, useState } from "react";

export default function BarcodeLabelGeneratorApp() {
  const [products, setProducts] = useState([]);
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState([]);
  const [loading, setLoading] = useState(false);

  const readExcel = async (file) => {
    setLoading(true);

    try {
      const XLSX = await import("xlsx");
      const reader = new FileReader();

      reader.onload = (evt) => {
        try {
          const data = evt.target.result;

          const workbook = XLSX.read(data, {
            type: "array",
          });

          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

          const rows = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            raw: false,
            defval: "",
          });

          const parsed = [];

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i] || [];

            const name = String(row[3] || "").trim();
            const barcode = String(row[10] || "").trim();

            if (name && barcode) {
              parsed.push({
                id: `${i}-${barcode}`,
                name,
                barcode,
              });
            }
          }

          setProducts(parsed);
        } catch (err) {
          console.error(err);
          alert("Hiba történt az Excel feldolgozásakor.");
        }

        setLoading(false);
      };

      reader.onerror = () => {
        alert("Nem sikerült megnyitni a fájlt.");
        setLoading(false);
      };

      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error(err);
      alert("Nem sikerült beolvasni az Excel fájlt.");
      setLoading(false);
    }
  };

  const filtered = useMemo(() => {
    if (!search) return products;

    const s = search.toLowerCase();

    return products.filter(
      (p) =>
        p.name.toLowerCase().includes(s) ||
        p.barcode.toLowerCase().includes(s)
    );
  }, [products, search]);

  const addItem = (item) => {
    setSelected((prev) => [...prev, item]);
  };

  const removeItem = (index) => {
    setSelected((prev) => prev.filter((_, i) => i !== index));
  };

  const clearAll = () => {
    setSelected([]);
  };

  const createBarcodeDataUrl = async (value) => {
    const JsBarcodeModule = await import("jsbarcode");
    const JsBarcode = JsBarcodeModule.default;

    const canvas = document.createElement("canvas");

    JsBarcode(canvas, value, {
      format: "CODE128",
      width: 1.8,
      height: 40,
      displayValue: false,
      margin: 0,
    });

    return canvas.toDataURL("image/png");
  };

  const exportWord = async () => {
    if (!selected.length) {
      alert("Nincs kiválasztott termék.");
      return;
    }

    try {
      setLoading(true);

      const docx = await import("docx");
      const fileSaverModule = await import("file-saver");

      const saveAs =
        fileSaverModule.saveAs ||
        fileSaverModule.default?.saveAs ||
        fileSaverModule.default;

      if (typeof saveAs !== "function") {
        throw new Error("A file-saver saveAs függvény nem érhető el.");
      }

      const {
        Document,
        Packer,
        Paragraph,
        Table,
        TableRow,
        TableCell,
        WidthType,
        AlignmentType,
        ImageRun,
        TextRun,
        VerticalAlign,
        BorderStyle,
        PageOrientation,
      } = docx;

      const labelsPerRow = 4;
      const rowsPerPage = 10;

      const labelWidthCm = 5.5;
      const labelHeightCm = 3;

      const cmToTwip = (cm) => Math.round(cm * 566.929);

      const tableRows = [];

      for (let r = 0; r < Math.ceil(selected.length / labelsPerRow); r++) {
        const rowCells = [];

        for (let c = 0; c < labelsPerRow; c++) {
          const index = r * labelsPerRow + c;
          const item = selected[index];

          if (!item) {
            rowCells.push(
              new TableCell({
                width: {
                  size: cmToTwip(labelWidthCm),
                  type: WidthType.DXA,
                },
                borders: {
                  top: { style: BorderStyle.NONE, size: 0 },
                  bottom: { style: BorderStyle.NONE, size: 0 },
                  left: { style: BorderStyle.NONE, size: 0 },
                  right: { style: BorderStyle.NONE, size: 0 },
                },
                children: [new Paragraph("")],
              })
            );

            continue;
          }

          const barcodeDataUrl = await createBarcodeDataUrl(item.barcode);
          const base64 = barcodeDataUrl.split(",")[1];

          const imageData = Uint8Array.from(atob(base64), (c) =>
            c.charCodeAt(0)
          );

          rowCells.push(
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: cmToTwip(labelWidthCm),
                type: WidthType.DXA,
              },
              margins: {
                top: 80,
                bottom: 80,
                left: 80,
                right: 80,
              },
              borders: {
                top: { style: BorderStyle.NONE, size: 0 },
                bottom: { style: BorderStyle.NONE, size: 0 },
                left: { style: BorderStyle.NONE, size: 0 },
                right: { style: BorderStyle.NONE, size: 0 },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: {
                    after: 60,
                  },
                  children: [
                    new TextRun({
                      text: item.name,
                      size: 16,
                      bold: true,
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: {
                    after: 60,
                  },
                  children: [
                    new ImageRun({
                      data: imageData,
                      transformation: {
                        width: 165,
                        height: 50,
                      },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: {
                    before: 20,
                  },
                  children: [
                    new TextRun({
                      text: item.barcode,
                      size: 16,
                    }),
                  ],
                }),
              ],
            })
          );
        }

        tableRows.push(
          new TableRow({
            height: {
              value: cmToTwip(labelHeightCm),
            },
            children: rowCells,
          })
        );
      }

      while (tableRows.length % rowsPerPage !== 0) {
        tableRows.push(
          new TableRow({
            height: {
              value: cmToTwip(labelHeightCm),
            },
            children: Array.from({ length: labelsPerRow }).map(
              () =>
                new TableCell({
                  borders: {
                    top: { style: BorderStyle.NONE, size: 0 },
                    bottom: { style: BorderStyle.NONE, size: 0 },
                    left: { style: BorderStyle.NONE, size: 0 },
                    right: { style: BorderStyle.NONE, size: 0 },
                  },
                  children: [new Paragraph("")],
                })
            ),
          })
        );
      }

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                size: {
                  orientation: PageOrientation.PORTRAIT,
                },
                margin: {
                  top: cmToTwip(0.19),
                  bottom: cmToTwip(0.19),
                  left: cmToTwip(0.19),
                  right: cmToTwip(0.19),
                },
              },
            },
            children: [
              new Table({
                width: {
                  size: 100,
                  type: WidthType.PERCENTAGE,
                },
                rows: tableRows,
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);

      saveAs(blob, "vonalkod_etikettek.docx");
    } catch (err) {
      console.error(err);
      alert(`Hiba történt a Word export során: ${err.message}`);
    }

    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <div className="max-w-7xl mx-auto grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-3xl shadow-xl p-6">
          <h1 className="text-3xl font-bold mb-2">
            Vonalkód etikett készítő
          </h1>

          <p className="text-gray-500 mb-6">
            Excel → keresés → kiválasztás → Word export
          </p>

          <div className="border-2 border-dashed rounded-2xl p-6 bg-gray-50 mb-6">
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  readExcel(file);
                }
              }}
              className="w-full"
            />

            <div className="mt-4 text-sm text-gray-600">
              D oszlop = termék neve
              <br />
              K oszlop = vonalkód
            </div>

            <div className="mt-3 font-semibold text-black">
              Betöltött termékek: {products.length} db
            </div>
          </div>

          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Keresés termékre vagy vonalkódra..."
            className="w-full border rounded-2xl px-4 py-3 mb-5 outline-none focus:ring-2 focus:ring-black"
          />

          <div className="h-[620px] overflow-y-auto border rounded-2xl divide-y bg-white">
            {filtered.length > 0 ? (
              filtered.map((item) => (
                <button
                  key={item.id}
                  onClick={() => addItem(item)}
                  className="w-full text-left p-4 hover:bg-gray-100 transition"
                >
                  <div className="font-semibold">{item.name}</div>
                  <div className="text-sm text-gray-500 mt-1">
                    {item.barcode}
                  </div>
                </button>
              ))
            ) : (
              <div className="p-8 text-center text-gray-400">
                Nincs betöltött vagy található termék.
              </div>
            )}
          </div>
        </div>

        <div className="bg-white rounded-3xl shadow-xl p-6 flex flex-col">
          <div className="flex items-center justify-between mb-5">
            <div>
              <h2 className="text-2xl font-bold">Kiválasztott címkék</h2>
              <div className="text-sm text-gray-500 mt-1">
                {selected.length} db etikett
              </div>
            </div>

            <div className="flex gap-3">
              <button
                onClick={clearAll}
                className="px-4 py-2 rounded-2xl border hover:bg-gray-100"
              >
                Törlés
              </button>

              <button
                onClick={exportWord}
                disabled={loading}
                className="px-5 py-2 rounded-2xl bg-black text-white disabled:opacity-50"
              >
                {loading ? "Feldolgozás..." : "Word DOC letöltés"}
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto border rounded-2xl divide-y bg-white">
            {selected.length > 0 ? (
              selected.map((item, index) => (
                <div
                  key={`${item.id}-${index}`}
                  className="p-4 flex items-center justify-between"
                >
                  <div>
                    <div className="font-semibold">{item.name}</div>
                    <div className="text-sm text-gray-500 mt-1">
                      {item.barcode}
                    </div>
                  </div>

                  <button
                    onClick={() => removeItem(index)}
                    className="px-3 py-2 border rounded-xl hover:bg-red-50"
                  >
                    ✕
                  </button>
                </div>
              ))
            ) : (
              <div className="h-full flex items-center justify-center text-gray-400 p-10 text-center">
                Válassz termékeket a bal oldali listából.
              </div>
            )}
          </div>

          <div className="mt-5 bg-gray-50 rounded-2xl p-4 text-sm text-gray-600">
            <div className="font-semibold mb-2">Nyomtatási beállítások</div>

            <ul className="list-disc pl-5 space-y-1">
              <li>A4 álló tájolás</li>
              <li>10 sor × 4 oszlop</li>
              <li>3 cm × 5.5 cm etikett</li>
              <li>0.19 cm margó</li>
              <li>Középre igazított vonalkód</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
