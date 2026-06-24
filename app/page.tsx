"use client";

import { ChangeEvent, FormEvent, useMemo, useRef, useState } from "react";
import { Grid2x2 } from "lucide-react";
import * as XLSX from "xlsx";

type RowData = {
  __id: string;
  [key: string]: string | number | boolean | null | undefined;
};

type EditMode = "add" | "edit" | null;

const acceptedExtensions = ["xlsx", "xls", "xlsm", "xlsb", "csv"];

function normalizeHeader(value: unknown, index: number) {
  const label = String(value ?? "").trim();
  return label || `Column ${index + 1}`;
}

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function stringifyCell(value: RowData[string]) {
  if (value === null || value === undefined) return "";
  return String(value);
}

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [fileName, setFileName] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [columns, setColumns] = useState<string[]>([]);
  const [rows, setRows] = useState<RowData[]>([]);
  const [query, setQuery] = useState("");
  const [error, setError] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);
  const [editMode, setEditMode] = useState<EditMode>(null);
  const [activeRowId, setActiveRowId] = useState<string | null>(null);
  const [formValues, setFormValues] = useState<Record<string, string>>({});

  const filteredRows = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return rows;

    return rows.filter((row) =>
      columns.some((column) => stringifyCell(row[column]).toLowerCase().includes(q))
    );
  }, [rows, columns, query]);

  const stats = useMemo(() => {
    return {
      totalRows: rows.length,
      totalColumns: columns.length,
      visibleRows: filteredRows.length,
    };
  }, [rows.length, columns.length, filteredRows.length]);

  async function handleFileUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    const extension = file.name.split(".").pop()?.toLowerCase() || "";
    if (!acceptedExtensions.includes(extension)) {
      setError("Please upload a valid Excel file (.xlsx, .xls, .xlsm, .xlsb) or CSV file.");
      setFileName("");
      setSheetName("");
      setColumns([]);
      setRows([]);
      return;
    }

    setError("");
    setIsProcessing(true);
    setFileName(file.name);
    setSheetName("");
    setColumns([]);
    setRows([]);
    setQuery("");

    try {
      await new Promise((resolve) => setTimeout(resolve, 1200));
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const firstSheetName = workbook.SheetNames[0];

      if (!firstSheetName) {
        throw new Error("No worksheet found in this file.");
      }

      const sheet = workbook.Sheets[firstSheetName];
      const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
        header: 1,
        defval: "",
        raw: false,
      });

      const nonEmptyRows = rawRows.filter((row) =>
        Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== "")
      );

      if (nonEmptyRows.length < 1) {
        throw new Error("The worksheet is empty. Please upload a file with headers and data.");
      }

      const headerRow = nonEmptyRows[0] as unknown[];
      const parsedColumns = headerRow.map((header, index) => normalizeHeader(header, index));
      const uniqueColumns = parsedColumns.map((column, index, arr) => {
        const countBefore = arr.slice(0, index).filter((item) => item === column).length;
        return countBefore > 0 ? `${column} ${countBefore + 1}` : column;
      });

      const parsedRows = nonEmptyRows.slice(1).map((row) => {
        const values = row as unknown[];
        const record: RowData = { __id: createId() };
        uniqueColumns.forEach((column, index) => {
          record[column] = values[index] === undefined ? "" : String(values[index]);
        });
        return record;
      });

      setSheetName(firstSheetName);
      setColumns(uniqueColumns);
      setRows(parsedRows);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unable to process this file.");
      setFileName("");
      setSheetName("");
      setColumns([]);
      setRows([]);
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  function openAddModal() {
    const emptyValues = Object.fromEntries(columns.map((column) => [column, ""]));
    setFormValues(emptyValues);
    setActiveRowId(null);
    setEditMode("add");
  }

  function openEditModal(row: RowData) {
    const values = Object.fromEntries(columns.map((column) => [column, stringifyCell(row[column])]));
    setFormValues(values);
    setActiveRowId(row.__id);
    setEditMode("edit");
  }

  function closeModal() {
    setEditMode(null);
    setActiveRowId(null);
    setFormValues({});
  }

  function submitForm(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (editMode === "add") {
      setRows((prev) => [{ __id: createId(), ...formValues }, ...prev]);
      closeModal();
      return;
    }

    if (editMode === "edit" && activeRowId) {
      setRows((prev) =>
        prev.map((row) => (row.__id === activeRowId ? { ...row, ...formValues } : row))
      );
      closeModal();
    }
  }

  function deleteRow(id: string) {
    const confirmed = window.confirm("Delete this record? This only removes it from the browser view.");
    if (!confirmed) return;
    setRows((prev) => prev.filter((row) => row.__id !== id));
  }

  function resetDemo() {
    setFileName("");
    setSheetName("");
    setColumns([]);
    setRows([]);
    setQuery("");
    setError("");
    closeModal();
  }

  return (
    <main className="min-h-screen bg-slate-50">
      <section className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-7xl flex-col gap-8 px-6 py-10 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <div className="mb-6 flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-xl border border-blue-200 bg-blue-50">
                <Grid2x2 className="h-6 w-6 text-blue-600" />
              </div>

              <span className="text-3xl font-black tracking-tight text-slate-950">
                SheetOS
              </span>
            </div>

            <div className="mb-8 inline-flex rounded-full border border-blue-200 bg-blue-50 px-4 py-2 text-sm font-semibold text-blue-700">
              Browser-only • No database • No upload to server
            </div>

            <h1 className="max-w-3xl text-4xl font-black tracking-tight text-slate-950 md:text-6xl">
              Turn any Excel sheet into a working app.
            </h1>

            <p className="mt-5 max-w-2xl text-lg leading-8 text-slate-600">
              Upload an Excel file and instantly generate a dynamic table with search, add,
              edit, and delete actions — fully inside the browser.
            </p>
          </div>
          <div className="rounded-3xl border border-slate-200 bg-slate-950 p-6 text-white shadow-soft lg:w-[360px]">
            <p className="text-sm font-semibold uppercase tracking-[0.2em] text-blue-300">Privacy first</p>
            <p className="mt-3 text-2xl font-bold">Files never leave your device.</p>
            <p className="mt-3 text-sm leading-6 text-slate-300">
              This prototype reads the spreadsheet in the browser and keeps everything in memory only.
            </p>
          </div>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-6 py-8">
        <div className="rounded-3xl border border-slate-200 bg-white p-6 shadow-soft">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <h2 className="text-2xl font-bold text-slate-950">Upload Excel file</h2>
              <p className="mt-1 text-sm text-slate-500">Supported: .xlsx, .xls, .xlsm, .xlsb, .csv</p>
            </div>
            <div className="flex flex-col gap-3 sm:flex-row">
              <label className="cursor-pointer rounded-2xl bg-blue-600 px-6 py-3 text-center font-bold text-white shadow-sm transition hover:bg-blue-700">
                Choose file
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
                  className="hidden"
                  onChange={handleFileUpload}
                />
              </label>
              {rows.length > 0 && (
                <button
                  onClick={resetDemo}
                  className="rounded-2xl border border-slate-300 bg-white px-6 py-3 font-bold text-slate-700 transition hover:bg-slate-50"
                >
                  Reset
                </button>
              )}
            </div>
          </div>

          {error && (
            <div className="mt-5 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-semibold text-red-700">
              {error}
            </div>
          )}

          {isProcessing && (
            <div className="mt-6 rounded-2xl border border-blue-200 bg-blue-50 p-5">
              <div className="flex items-center gap-4">
                <div className="h-10 w-10 animate-spin rounded-full border-4 border-blue-200 border-t-blue-600" />
                <div>
                  <p className="font-bold text-blue-950">Processing spreadsheet...</p>
                  <p className="text-sm text-blue-700">Reading columns, rows, and preparing your dynamic CRUD interface.</p>
                </div>
              </div>
            </div>
          )}

          {rows.length > 0 && !isProcessing && (
            <div className="mt-6 grid gap-4 md:grid-cols-4">
              <StatCard label="File" value={fileName} />
              <StatCard label="Sheet" value={sheetName} />
              <StatCard label="Columns" value={stats.totalColumns.toString()} />
              <StatCard label="Rows" value={stats.totalRows.toString()} />
            </div>
          )}
        </div>
      </section>

      {rows.length > 0 && !isProcessing && (
        <section className="mx-auto max-w-7xl px-6 pb-12">
          <div className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-soft">
            <div className="flex flex-col gap-4 border-b border-slate-200 p-5 lg:flex-row lg:items-center lg:justify-between">
              <div>
                <h2 className="text-2xl font-bold text-slate-950">Generated CRUD Table</h2>
                <p className="text-sm text-slate-500">
                  Showing {stats.visibleRows} of {stats.totalRows} records
                </p>
              </div>
              <div className="flex flex-col gap-3 sm:flex-row">
                <input
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="Search anything..."
                  className="w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none ring-blue-200 transition focus:ring-4 sm:w-72"
                />
                <button
                  onClick={openAddModal}
                  className="rounded-2xl bg-slate-950 px-6 py-3 font-bold text-white transition hover:bg-slate-800"
                >
                  + Add New
                </button>
              </div>
            </div>

            <div className="max-h-[620px] overflow-auto">
              <table className="min-w-full border-collapse text-left text-sm">
                <thead className="sticky top-0 z-10 bg-slate-100 text-xs uppercase tracking-wide text-slate-600">
                  <tr>
                    {columns.map((column) => (
                      <th key={column} className="whitespace-nowrap border-b border-slate-200 px-4 py-3 font-black">
                        {column}
                      </th>
                    ))}
                    <th className="sticky right-0 whitespace-nowrap border-b border-slate-200 bg-slate-100 px-4 py-3 font-black">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((row) => (
                    <tr key={row.__id} className="border-b border-slate-100 hover:bg-blue-50/40">
                      {columns.map((column) => (
                        <td key={`${row.__id}-${column}`} className="max-w-[260px] truncate whitespace-nowrap px-4 py-3 text-slate-700">
                          {stringifyCell(row[column]) || <span className="text-slate-300">—</span>}
                        </td>
                      ))}
                      <td className="sticky right-0 whitespace-nowrap bg-white px-4 py-3 shadow-[-12px_0_20px_rgba(255,255,255,0.8)]">
                        <div className="flex gap-2">
                          <button
                            onClick={() => openEditModal(row)}
                            className="rounded-xl border border-slate-300 px-3 py-2 text-xs font-bold text-slate-700 hover:bg-slate-50"
                          >
                            Edit
                          </button>
                          <button
                            onClick={() => deleteRow(row.__id)}
                            className="rounded-xl border border-red-200 bg-red-50 px-3 py-2 text-xs font-bold text-red-700 hover:bg-red-100"
                          >
                            Delete
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>

              {filteredRows.length === 0 && (
                <div className="p-12 text-center">
                  <p className="text-lg font-bold text-slate-700">No matching records found.</p>
                  <p className="mt-1 text-sm text-slate-500">Try a different search term.</p>
                </div>
              )}
            </div>
          </div>
        </section>
      )}

      {editMode && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4 backdrop-blur-sm">
          <div className="max-h-[90vh] w-full max-w-3xl overflow-hidden rounded-3xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-200 p-5">
              <div>
                <h3 className="text-2xl font-black text-slate-950">{editMode === "add" ? "Add New Record" : "Edit Record"}</h3>
                <p className="text-sm text-slate-500">Fields are generated dynamically from your Excel columns.</p>
              </div>
              <button onClick={closeModal} className="rounded-full p-2 text-slate-500 hover:bg-slate-100">
                ✕
              </button>
            </div>

            <form onSubmit={submitForm}>
              <div className="grid max-h-[60vh] gap-4 overflow-auto p-5 md:grid-cols-2">
                {columns.map((column) => (
                  <label key={column} className="block">
                    <span className="mb-2 block text-sm font-bold text-slate-700">{column}</span>
                    <input
                      value={formValues[column] ?? ""}
                      onChange={(event) => setFormValues((prev) => ({ ...prev, [column]: event.target.value }))}
                      className="w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none ring-blue-200 transition focus:ring-4"
                    />
                  </label>
                ))}
              </div>

              <div className="flex justify-end gap-3 border-t border-slate-200 p-5">
                <button
                  type="button"
                  onClick={closeModal}
                  className="rounded-2xl border border-slate-300 px-5 py-3 font-bold text-slate-700 hover:bg-slate-50"
                >
                  Cancel
                </button>
                <button type="submit" className="rounded-2xl bg-blue-600 px-5 py-3 font-bold text-white hover:bg-blue-700">
                  {editMode === "add" ? "Create Record" : "Save Changes"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      <section className="mx-auto mt-14 mb-14 max-w-7xl px-6">
        <div className="rounded-[32px] border border-slate-200 bg-white p-8 shadow-sm">
          <div className="mb-8 flex items-end justify-between gap-4">
            <div>
              <p className="text-sm font-semibold uppercase tracking-[0.25em] text-blue-600">
                What happens next
              </p>
              <h2 className="mt-3 text-3xl font-black tracking-tight text-slate-950">
                Your excel sheet becomes a usable app instantly.
              </h2>
            </div>
            <p className="max-w-md text-sm text-slate-500">
              SheetOS reads your columns, detects the structure, and generates a browser-only workspace.
            </p>
          </div>

          <div className="grid gap-4 md:grid-cols-4">
            {[
              ["01", "Upload Excel", "Drop any .xlsx, .xls or .csv file."],
              ["02", "Detect Columns", "Headers become dynamic app fields."],
              ["03", "Generate App", "Search, add, edit and delete records."],
              ["04", "Stay Private", "Everything stays inside your browser."]
            ].map(([step, title, desc]) => (
              <div
                key={step}
                className="rounded-2xl border border-slate-200 bg-slate-50 p-5"
              >
                <div className="mb-8 flex h-10 w-10 items-center justify-center rounded-full bg-blue-600 text-sm font-bold text-white">
                  {step}
                </div>
                <h3 className="text-lg font-bold text-slate-950">{title}</h3>
                <p className="mt-2 text-sm leading-6 text-slate-600">{desc}</p>
              </div>
            ))}
          </div>
        </div>
      </section>
    </main>
  );
}

function StatCard({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
      <p className="text-xs font-bold uppercase tracking-wide text-slate-500">{label}</p>
      <p className="mt-2 truncate text-lg font-black text-slate-950" title={value}>
        {value}
      </p>
    </div>
  );
}
