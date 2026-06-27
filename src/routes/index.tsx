import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, CartesianGrid, Legend,
} from "recharts";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Switch } from "@/components/ui/switch";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Separator } from "@/components/ui/separator";
import {
  Upload, Sparkles, Database, Trash2, Download, FileSpreadsheet,
  FileText, BarChart3, CheckCircle2, AlertCircle, Github,
} from "lucide-react";
import {
  cleanData, defaultOptions, generateMessySample,
  type Row, type CleanReport, type CleanOptions,
} from "@/lib/data-cleaner";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "DataPolish — Automated Data Cleaning & Reporting" },
      { name: "description", content: "Upload messy CSV or Excel data, auto-clean missing values and duplicates, and export polished dashboards and reports." },
      { property: "og:title", content: "DataPolish — Automated Data Cleaning & Reporting" },
      { property: "og:description", content: "Upload messy CSV or Excel data, auto-clean missing values and duplicates, and export polished dashboards and reports." },
    ],
  }),
  component: Home,
});

const CHART_COLORS = ["var(--chart-1)", "var(--chart-2)", "var(--chart-3)", "var(--chart-4)", "var(--chart-5)"];

function Home() {
  const [rawRows, setRawRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [options, setOptions] = useState<CleanOptions>(defaultOptions);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const { cleaned, report } = useMemo(
    () => (rawRows.length ? cleanData(rawRows, options) : { cleaned: [], report: null as CleanReport | null }),
    [rawRows, options],
  );

  const handleFile = (file: File) => {
    setFileName(file.name);
    const ext = file.name.split(".").pop()?.toLowerCase();
    const reader = new FileReader();
    reader.onload = (e) => {
      const result = e.target?.result;
      if (!result) return;
      if (ext === "csv" || ext === "txt") {
        Papa.parse(result as string, {
          header: true,
          skipEmptyLines: true,
          complete: (out) => setRawRows(out.data as Row[]),
        });
      } else {
        const wb = XLSX.read(result as ArrayBuffer, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Row>(sheet, { defval: null, raw: false });
        setRawRows(json);
      }
    };
    if (ext === "csv" || ext === "txt") reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  };

  const loadSample = () => {
    setFileName("sample_messy_sales.csv");
    setRawRows(generateMessySample());
  };

  const clearAll = () => {
    setRawRows([]); setFileName("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const downloadExcel = () => {
    if (!cleaned.length || !report) return;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cleaned), "Cleaned Data");

    const summary = [
      ["Metric", "Value"],
      ["Original rows", report.originalRows],
      ["Cleaned rows", report.cleanedRows],
      ["Duplicates removed", report.duplicatesRemoved],
      ["Missing values filled", report.missingFilled],
      ["Whitespace trimmed", report.whitespaceTrimmed],
      ["Case normalized", report.caseNormalized],
      ["Dates normalized", report.datesNormalized],
      ["Numbers coerced", report.numbersCoerced],
      ["Generated at", new Date().toISOString()],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), "Summary");

    const colsSheet = [
      ["Column", "Type", "Filled", "Missing", "Unique", "Min", "Max", "Mean"],
      ...report.columns.map((c) => [
        c.name, c.type, c.filled, c.missing, c.unique,
        c.min ?? "", c.max ?? "", c.mean !== undefined ? Number(c.mean.toFixed(2)) : "",
      ]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(colsSheet), "Column Stats");

    XLSX.writeFile(wb, `${fileName.replace(/\.[^.]+$/, "") || "report"}_cleaned.xlsx`);
  };

  const downloadPdf = () => {
    if (!cleaned.length || !report) return;
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const margin = 40;

    doc.setFontSize(20); doc.text("Data Cleaning Report", margin, 50);
    doc.setFontSize(10); doc.setTextColor(100);
    doc.text(`Source: ${fileName || "uploaded data"}`, margin, 68);
    doc.text(`Generated: ${new Date().toLocaleString()}`, margin, 82);

    autoTable(doc, {
      startY: 100,
      head: [["Metric", "Value"]],
      body: [
        ["Original rows", String(report.originalRows)],
        ["Cleaned rows", String(report.cleanedRows)],
        ["Duplicates removed", String(report.duplicatesRemoved)],
        ["Missing values filled", String(report.missingFilled)],
        ["Whitespace trimmed", String(report.whitespaceTrimmed)],
        ["Case normalized", String(report.caseNormalized)],
        ["Dates normalized", String(report.datesNormalized)],
        ["Numbers coerced", String(report.numbersCoerced)],
      ],
      theme: "striped",
      headStyles: { fillColor: [82, 82, 220] },
    });

    autoTable(doc, {
      head: [["Column", "Type", "Filled", "Missing", "Unique", "Min", "Max", "Mean"]],
      body: report.columns.map((c) => [
        c.name, c.type, c.filled, c.missing, c.unique,
        c.min ?? "-", c.max ?? "-", c.mean !== undefined ? c.mean.toFixed(2) : "-",
      ]),
      theme: "grid",
      headStyles: { fillColor: [82, 82, 220] },
      styles: { fontSize: 9 },
    });

    autoTable(doc, {
      head: [Object.keys(cleaned[0] ?? {})],
      body: cleaned.slice(0, 25).map((r) => Object.values(r).map((v) => (v == null ? "" : String(v)))),
      theme: "striped",
      headStyles: { fillColor: [82, 82, 220] },
      styles: { fontSize: 8, cellPadding: 3 },
    });

    doc.save(`${fileName.replace(/\.[^.]+$/, "") || "report"}_report.pdf`);
  };

  const downloadCleanedCsv = () => {
    if (!cleaned.length) return;
    const csv = Papa.unparse(cleaned);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${fileName.replace(/\.[^.]+$/, "") || "data"}_cleaned.csv`;
    a.click(); URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b bg-card/50 backdrop-blur sticky top-0 z-10">
        <div className="mx-auto max-w-7xl px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="size-9 rounded-lg bg-primary text-primary-foreground grid place-items-center">
              <Sparkles className="size-5" />
            </div>
            <div>
              <h1 className="text-lg font-semibold leading-none">DataPolish</h1>
              <p className="text-xs text-muted-foreground">Automated cleaning & reporting</p>
            </div>
          </div>
          <a
            href="https://github.com"
            target="_blank" rel="noreferrer"
            className="text-sm text-muted-foreground hover:text-foreground inline-flex items-center gap-2"
          >
            <Github className="size-4" /> View on GitHub
          </a>
        </div>
      </header>

      <main className="mx-auto max-w-7xl px-6 py-8 space-y-8">
        {!rawRows.length ? (
          <Hero onSample={loadSample} onUpload={() => fileInputRef.current?.click()} />
        ) : (
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-2xl font-semibold tracking-tight">{fileName}</h2>
              <p className="text-sm text-muted-foreground">
                {report?.originalRows} rows in · {report?.cleanedRows} rows out
              </p>
            </div>
            <div className="flex gap-2">
              <Button variant="outline" onClick={clearAll}><Trash2 className="size-4" /> Clear</Button>
              <Button variant="outline" onClick={() => fileInputRef.current?.click()}>
                <Upload className="size-4" /> Replace
              </Button>
            </div>
          </div>
        )}

        <input
          ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.txt" className="hidden"
          onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
        />

        {report && (
          <>
            <SummaryRow report={report} />
            <div className="grid gap-6 lg:grid-cols-[280px_1fr]">
              <OptionsPanel options={options} setOptions={setOptions} />
              <Tabs defaultValue="dashboard">
                <TabsList>
                  <TabsTrigger value="dashboard"><BarChart3 className="size-4" /> Dashboard</TabsTrigger>
                  <TabsTrigger value="preview"><Database className="size-4" /> Data preview</TabsTrigger>
                  <TabsTrigger value="columns"><FileText className="size-4" /> Columns</TabsTrigger>
                  <TabsTrigger value="export"><Download className="size-4" /> Export</TabsTrigger>
                </TabsList>
                <TabsContent value="dashboard" className="mt-4">
                  <Dashboard cleaned={cleaned} report={report} />
                </TabsContent>
                <TabsContent value="preview" className="mt-4">
                  <DataPreview rows={cleaned} />
                </TabsContent>
                <TabsContent value="columns" className="mt-4">
                  <ColumnStatsTable report={report} />
                </TabsContent>
                <TabsContent value="export" className="mt-4">
                  <ExportPanel
                    onExcel={downloadExcel} onPdf={downloadPdf} onCsv={downloadCleanedCsv}
                  />
                </TabsContent>
              </Tabs>
            </div>
          </>
        )}

        <Footer />
      </main>
    </div>
  );
}

function Hero({ onSample, onUpload }: { onSample: () => void; onUpload: () => void }) {
  return (
    <section className="text-center py-16 space-y-6">
      <Badge variant="secondary" className="mx-auto">Python · Excel · Power BI workflows, in your browser</Badge>
      <h2 className="text-5xl font-semibold tracking-tight max-w-3xl mx-auto">
        Clean messy data and ship a report in one click.
      </h2>
      <p className="text-muted-foreground max-w-xl mx-auto text-lg">
        Upload a CSV or Excel file. We detect column types, fill missing values, drop duplicates,
        normalize formats, and generate a dashboard + downloadable report.
      </p>
      <div className="flex flex-wrap justify-center gap-3 pt-2">
        <Button size="lg" onClick={onUpload}><Upload className="size-4" /> Upload CSV / Excel</Button>
        <Button size="lg" variant="outline" onClick={onSample}>
          <Sparkles className="size-4" /> Try sample dataset
        </Button>
      </div>
      <div className="grid sm:grid-cols-3 gap-4 max-w-4xl mx-auto pt-12 text-left">
        {[
          { icon: CheckCircle2, title: "Smart cleaning", body: "Detects numeric, date, and string columns; fills missing values with mean/median/mode." },
          { icon: Database, title: "Duplicate handling", body: "Removes exact duplicates after normalization, so 'North' and 'north ' count as one." },
          { icon: FileSpreadsheet, title: "Multi-format export", body: "Cleaned CSV, multi-sheet Excel workbook, or printable PDF report — your choice." },
        ].map((f) => (
          <Card key={f.title}>
            <CardHeader>
              <f.icon className="size-5 text-primary" />
              <CardTitle className="text-base">{f.title}</CardTitle>
              <CardDescription>{f.body}</CardDescription>
            </CardHeader>
          </Card>
        ))}
      </div>
    </section>
  );
}

function SummaryRow({ report }: { report: CleanReport }) {
  const stats = [
    { label: "Rows in", value: report.originalRows, tone: "muted" },
    { label: "Rows out", value: report.cleanedRows, tone: "primary" },
    { label: "Duplicates removed", value: report.duplicatesRemoved, tone: "warning" },
    { label: "Missing filled", value: report.missingFilled, tone: "success" },
    { label: "Whitespace trimmed", value: report.whitespaceTrimmed, tone: "muted" },
    { label: "Dates normalized", value: report.datesNormalized, tone: "muted" },
  ];
  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
      {stats.map((s) => (
        <Card key={s.label}>
          <CardContent className="py-4">
            <div className="text-2xl font-semibold">{s.value.toLocaleString()}</div>
            <div className="text-xs text-muted-foreground mt-1">{s.label}</div>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}

function OptionsPanel({
  options, setOptions,
}: { options: CleanOptions; setOptions: (o: CleanOptions) => void }) {
  const set = <K extends keyof CleanOptions>(k: K, v: CleanOptions[K]) =>
    setOptions({ ...options, [k]: v });

  return (
    <Card className="h-fit sticky top-24">
      <CardHeader>
        <CardTitle className="text-base">Cleaning rules</CardTitle>
        <CardDescription>Re-runs instantly as you toggle.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <ToggleRow label="Drop duplicates" checked={options.dropDuplicates}
          onChange={(v) => set("dropDuplicates", v)} />
        <ToggleRow label="Trim whitespace" checked={options.trimWhitespace}
          onChange={(v) => set("trimWhitespace", v)} />
        <ToggleRow label="Title-case categories" checked={options.normalizeCase}
          onChange={(v) => set("normalizeCase", v)} />
        <ToggleRow label="Normalize dates (ISO)" checked={options.normalizeDates}
          onChange={(v) => set("normalizeDates", v)} />
        <ToggleRow label="Drop empty columns" checked={options.dropEmptyColumns}
          onChange={(v) => set("dropEmptyColumns", v)} />
        <Separator />
        <div className="space-y-2">
          <Label className="text-xs">Fill missing numbers with</Label>
          <Select value={options.fillNumericWith}
            onValueChange={(v) => set("fillNumericWith", v as CleanOptions["fillNumericWith"])}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="median">Median</SelectItem>
              <SelectItem value="mean">Mean</SelectItem>
              <SelectItem value="zero">Zero</SelectItem>
              <SelectItem value="none">Leave empty</SelectItem>
            </SelectContent>
          </Select>
        </div>
        <div className="space-y-2">
          <Label className="text-xs">Fill missing text with</Label>
          <Select value={options.fillStringWith}
            onValueChange={(v) => set("fillStringWith", v as CleanOptions["fillStringWith"])}>
            <SelectTrigger><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="mode">Most common</SelectItem>
              <SelectItem value="unknown">"Unknown"</SelectItem>
              <SelectItem value="none">Leave empty</SelectItem>
            </SelectContent>
          </Select>
        </div>
      </CardContent>
    </Card>
  );
}

function ToggleRow({
  label, checked, onChange,
}: { label: string; checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <div className="flex items-center justify-between">
      <Label className="text-sm font-normal">{label}</Label>
      <Switch checked={checked} onCheckedChange={onChange} />
    </div>
  );
}

function Dashboard({ cleaned, report }: { cleaned: Row[]; report: CleanReport }) {
  const numericCols = report.columns.filter((c) => c.type === "number");
  const stringCols = report.columns.filter((c) => c.type === "string");
  const dateCols = report.columns.filter((c) => c.type === "date");

  const primaryNumeric = numericCols[0];
  const primaryCategory = stringCols.find((c) => c.unique > 1 && c.unique <= 12) ?? stringCols[0];
  const dateCol = dateCols[0];

  // Bar: category counts
  const categoryData = useMemo(() => {
    if (!primaryCategory) return [];
    const counts: Record<string, number> = {};
    for (const r of cleaned) {
      const k = String(r[primaryCategory.name] ?? "—");
      counts[k] = (counts[k] ?? 0) + 1;
    }
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 8)
      .map(([name, value]) => ({ name, value }));
  }, [cleaned, primaryCategory]);

  // Pie: sum of numeric by category
  const pieData = useMemo(() => {
    if (!primaryCategory || !primaryNumeric) return [];
    const sums: Record<string, number> = {};
    for (const r of cleaned) {
      const k = String(r[primaryCategory.name] ?? "—");
      const v = Number(r[primaryNumeric.name]);
      if (Number.isFinite(v)) sums[k] = (sums[k] ?? 0) + v;
    }
    return Object.entries(sums).sort((a, b) => b[1] - a[1]).slice(0, 6)
      .map(([name, value]) => ({ name, value }));
  }, [cleaned, primaryCategory, primaryNumeric]);

  // Line: numeric over time
  const timeData = useMemo(() => {
    if (!dateCol || !primaryNumeric) return [];
    const map = new Map<string, number>();
    for (const r of cleaned) {
      const d = String(r[dateCol.name] ?? "").slice(0, 7); // YYYY-MM
      if (!d) continue;
      const v = Number(r[primaryNumeric.name]);
      if (Number.isFinite(v)) map.set(d, (map.get(d) ?? 0) + v);
    }
    return [...map.entries()].sort(([a], [b]) => a.localeCompare(b))
      .map(([name, value]) => ({ name, value }));
  }, [cleaned, dateCol, primaryNumeric]);

  return (
    <div className="grid gap-4 md:grid-cols-2">
      {categoryData.length > 0 && (
        <ChartCard title={`Records by ${primaryCategory!.name}`}>
          <BarChart data={categoryData}>
            <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
            <XAxis dataKey="name" fontSize={11} />
            <YAxis fontSize={11} />
            <Tooltip />
            <Bar dataKey="value" fill="var(--chart-1)" radius={[6, 6, 0, 0]} />
          </BarChart>
        </ChartCard>
      )}
      {pieData.length > 0 && (
        <ChartCard title={`${primaryNumeric!.name} by ${primaryCategory!.name}`}>
          <PieChart>
            <Pie data={pieData} dataKey="value" nameKey="name" outerRadius={100} label>
              {pieData.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
            </Pie>
            <Tooltip />
            <Legend />
          </PieChart>
        </ChartCard>
      )}
      {timeData.length > 1 && (
        <ChartCard title={`${primaryNumeric!.name} over time`} wide>
          <LineChart data={timeData}>
            <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
            <XAxis dataKey="name" fontSize={11} />
            <YAxis fontSize={11} />
            <Tooltip />
            <Line type="monotone" dataKey="value" stroke="var(--chart-1)" strokeWidth={2} dot={false} />
          </LineChart>
        </ChartCard>
      )}
      {numericCols.length > 0 && (
        <ChartCard title="Numeric column ranges" wide={timeData.length <= 1}>
          <BarChart data={numericCols.map((c) => ({ name: c.name, min: c.min ?? 0, max: c.max ?? 0, mean: c.mean ?? 0 }))}>
            <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
            <XAxis dataKey="name" fontSize={11} />
            <YAxis fontSize={11} />
            <Tooltip />
            <Legend />
            <Bar dataKey="min" fill="var(--chart-2)" />
            <Bar dataKey="mean" fill="var(--chart-1)" />
            <Bar dataKey="max" fill="var(--chart-3)" />
          </BarChart>
        </ChartCard>
      )}
    </div>
  );
}

function ChartCard({ title, children, wide }: { title: string; children: React.ReactElement; wide?: boolean }) {
  return (
    <Card className={wide ? "md:col-span-2" : ""}>
      <CardHeader><CardTitle className="text-sm font-medium">{title}</CardTitle></CardHeader>
      <CardContent>
        <div className="h-72">
          <ResponsiveContainer width="100%" height="100%">{children}</ResponsiveContainer>
        </div>
      </CardContent>
    </Card>
  );
}

function DataPreview({ rows }: { rows: Row[] }) {
  const preview = rows.slice(0, 50);
  const cols = preview.length ? Object.keys(preview[0]) : [];
  return (
    <Card>
      <CardHeader>
        <CardTitle className="text-sm">First 50 cleaned rows</CardTitle>
        <CardDescription>Showing {preview.length} of {rows.length.toLocaleString()}</CardDescription>
      </CardHeader>
      <CardContent>
        <div className="overflow-auto max-h-[520px] rounded-md border">
          <Table>
            <TableHeader>
              <TableRow>{cols.map((c) => <TableHead key={c}>{c}</TableHead>)}</TableRow>
            </TableHeader>
            <TableBody>
              {preview.map((r, i) => (
                <TableRow key={i}>
                  {cols.map((c) => (
                    <TableCell key={c} className="font-mono text-xs">
                      {r[c] == null ? <span className="text-muted-foreground italic">null</span> : String(r[c])}
                    </TableCell>
                  ))}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      </CardContent>
    </Card>
  );
}

function ColumnStatsTable({ report }: { report: CleanReport }) {
  return (
    <Card>
      <CardHeader><CardTitle className="text-sm">Column statistics</CardTitle></CardHeader>
      <CardContent>
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Column</TableHead><TableHead>Type</TableHead>
              <TableHead className="text-right">Filled</TableHead>
              <TableHead className="text-right">Missing</TableHead>
              <TableHead className="text-right">Unique</TableHead>
              <TableHead className="text-right">Min</TableHead>
              <TableHead className="text-right">Max</TableHead>
              <TableHead className="text-right">Mean</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {report.columns.map((c) => (
              <TableRow key={c.name}>
                <TableCell className="font-medium">{c.name}</TableCell>
                <TableCell><Badge variant="outline">{c.type}</Badge></TableCell>
                <TableCell className="text-right">{c.filled}</TableCell>
                <TableCell className="text-right">
                  {c.missing > 0 ? <span className="text-warning inline-flex items-center gap-1"><AlertCircle className="size-3" />{c.missing}</span> : c.missing}
                </TableCell>
                <TableCell className="text-right">{c.unique}</TableCell>
                <TableCell className="text-right">{c.min ?? "—"}</TableCell>
                <TableCell className="text-right">{c.max ?? "—"}</TableCell>
                <TableCell className="text-right">{c.mean !== undefined ? c.mean.toFixed(2) : "—"}</TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  );
}

function ExportPanel({
  onExcel, onPdf, onCsv,
}: { onExcel: () => void; onPdf: () => void; onCsv: () => void }) {
  const exports = [
    { icon: FileSpreadsheet, title: "Excel workbook", body: "Cleaned data, summary, and column stats across three sheets.", action: onExcel, cta: "Download .xlsx" },
    { icon: FileText, title: "PDF report", body: "Printable report with summary metrics, column stats, and sample rows.", action: onPdf, cta: "Download .pdf" },
    { icon: Database, title: "Cleaned CSV", body: "Just the cleaned rows — drop into Power BI, Tableau, or pandas.", action: onCsv, cta: "Download .csv" },
  ];
  return (
    <div className="grid md:grid-cols-3 gap-4">
      {exports.map((e) => (
        <Card key={e.title}>
          <CardHeader>
            <e.icon className="size-5 text-primary" />
            <CardTitle className="text-base">{e.title}</CardTitle>
            <CardDescription>{e.body}</CardDescription>
          </CardHeader>
          <CardContent>
            <Button onClick={e.action} className="w-full"><Download className="size-4" /> {e.cta}</Button>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}

function Footer() {
  return (
    <footer className="border-t pt-8 mt-16 text-sm text-muted-foreground">
      <div className="flex flex-wrap justify-between gap-4">
        <p>Built with TanStack Start, Recharts, SheetJS, and jsPDF. Companion Python script in <code className="text-foreground">/python</code>.</p>
        <p>© {new Date().getFullYear()} DataPolish</p>
      </div>
    </footer>
  );
}
