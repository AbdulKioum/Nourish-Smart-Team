import  { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

function calcAgeDays(hatchDate?: string, visitDate?: string): number {
  if (!hatchDate) return 0;
  const hd = new Date(hatchDate);
  const vd = visitDate ? new Date(visitDate) : new Date();
  return Math.max(0, Math.floor((vd.getTime() - hd.getTime()) / 86400000));
}

export default function ReportsPage() {
  const [selectedCategory, setSelectedCategory] = useState<
    "Broiler" | "Layer" | "Sonali"
  >("Broiler");
  const [reportType, setReportType] = useState<"before_sale" | "after_sale">(
    "before_sale"
  );

  const [agents, setAgents] = useState<any[]>([]);
  const [farmers, setFarmers] = useState<any[]>([]);
  const [visitsMap, setVisitsMap] = useState<Record<string, any>>({});
  const [beforeReports, setBeforeReports] = useState<any[]>([]);
  const [afterReports, setAfterReports] = useState<any[]>([]);
  const [layerReports, setLayerReports] = useState<any[]>([]);
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>({});
  const [layerStds, setLayerStds] = useState<Record<number, any>>({});
  const [loading, setLoading] = useState(true);

  // Filters
  const [agentFilter, setAgentFilter] = useState<string>("all");
  const [farmerFilter, setFarmerFilter] = useState<string>("all");
  const [monthFilter, setMonthFilter] = useState<string>("");
  const [fromDate, setFromDate] = useState<string>("");
  const [toDate, setToDate] = useState<string>("");

  useEffect(() => {
    loadReportsData();
  }, []);

  async function loadReportsData() {
    setLoading(true);
    const [
      { data: agData },
      { data: frmData },
      { data: vData },
      { data: bRep },
      { data: aRep },
      { data: lRep },
      { data: bStd },
      { data: lStd },
    ] = await Promise.all([
      supabase.from("agents").select("*").order("agent_name"),
      supabase.from("farmers").select("*").order("farmer_name"),
      supabase.from("vet_visits").select("*"),
      supabase
        .from("before_sale_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase
        .from("after_sale_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase
        .from("layer_reports")
        .select("*")
        .order("visit_date", { ascending: false }),
      supabase.from("broiler_standards").select("*"),
      supabase.from("layer_standards").select("*"),
    ]);

    setAgents(agData || []);
    setFarmers(frmData || []);

    const vMap: Record<string, any> = {};
    (vData || []).forEach((v: any) => (vMap[String(v.id)] = v));
    setVisitsMap(vMap);

    setBeforeReports(bRep || []);
    setAfterReports(aRep || []);
    setLayerReports(lRep || []);

    const bMap: Record<number, any> = {};
    (bStd || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
    setBroilerStds(bMap);

    const lMap: Record<number, any> = {};
    (lStd || []).forEach((s: any) => (lMap[Number(s.age_weeks)] = s));
    setLayerStds(lMap);

    setLoading(false);
  }

  function handleMonthSelect(ym: string) {
    setMonthFilter(ym);
    if (!ym) {
      setFromDate("");
      setToDate("");
      return;
    }
    const [y, m] = ym.split("-").map(Number);
    const lastDay = new Date(y, m, 0).getDate();
    setFromDate(`${ym}-01`);
    setToDate(`${ym}-${String(lastDay).padStart(2, "0")}`);
  }

  const availableFarmers =
    agentFilter === "all"
      ? farmers
      : farmers.filter((f) => String(f.agent_id) === String(agentFilter));

  const rawActiveReports =
    selectedCategory === "Layer"
      ? layerReports
      : reportType === "before_sale"
      ? beforeReports
      : afterReports;

  // Enrich and Filter Reports
  const filteredReports = rawActiveReports
    .map((r) => {
      const visit = visitsMap[String(r.vet_visit_id)] || {};
      const farmer =
        farmers.find((f) => String(f.id) === String(r.farmer_id)) || {};
      const agent =
        agents.find((a) => String(a.id) === String(r.agent_id)) || {};
      const cat =
        visit.category ||
        farmer.farm_type ||
        (selectedCategory === "Layer" ? "Layer" : "Broiler");

      return {
        ...r,
        visit,
        farmer,
        agent,
        category: cat,
        batchKey: String(r.vet_visit_id || r.farmer_id),
        batchTitle:
          visit.batch_label || farmer.farmer_name || "Unknown Farmer",
      };
    })
    .filter((item) => {
      if (
        selectedCategory !== "Layer" &&
        item.category?.toLowerCase() !== selectedCategory.toLowerCase()
      ) {
        return false;
      }
      if (
        agentFilter !== "all" &&
        String(item.agent_id) !== String(agentFilter)
      ) {
        return false;
      }
      if (
        farmerFilter !== "all" &&
        String(item.farmer_id) !== String(farmerFilter)
      ) {
        return false;
      }
      if (fromDate && item.visit_date < fromDate) return false;
      if (toDate && item.visit_date > toDate) return false;
      return true;
    });

  // Collect Available Months from all reports
  const allMonths = Array.from(
    new Set(
      [...beforeReports, ...afterReports, ...layerReports]
        .map((r) => String(r.visit_date || "").slice(0, 7))
        .filter((m) => m.length === 7)
    )
  )
    .sort()
    .reverse();

  // Group by Batch (`vet_visit_id`) so filtering a farmer shows each batch cleanly
  const groupedBatches = Object.values(
    filteredReports.reduce((acc: Record<string, any>, rep: any) => {
      const key = rep.batchKey;
      if (!acc[key]) {
        acc[key] = {
          batchKey: key,
          batchTitle: rep.batchTitle,
          agentName: rep.agent?.agent_name || "-",
          farmerName: rep.farmer?.farmer_name || "-",
          hatchDate: rep.visit?.hatch_date || "-",
          breed: rep.visit?.breed || "-",
          tBirds: rep.visit?.t_birds || 0,
          status: rep.visit?.status || "-",
          reports: [],
        };
      }
      acc[key].reports.push(rep);
      return acc;
    }, {})
  );

  async function exportFilteredReportsExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(`${selectedCategory}_Reports`);

    if (selectedCategory === "Layer") {
      sheet.addRow([
        "Visit Date",
        "Agent",
        "Farmer / Batch",
        "H.Date",
        "T.Birds",
        "Age (Wks)",
        "Weight (g)",
        "Std Wt",
        "Feed (g/b/d)",
        "Std Feed",
        "Egg Prod (%)",
        "Std Egg%",
        "Egg Wt (g)",
        "Std Egg Wt",
        "Feed Company",
        "Feed Type",
        "Remarks",
      ]);

      filteredReports.forEach((r: any) => {
        const wks = Math.max(
          1,
          Math.floor(calcAgeDays(r.visit?.hatch_date, r.visit_date) / 7)
        );
        const std = layerStds[wks];
        sheet.addRow([
          r.visit_date,
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.visit?.hatch_date || "-",
          r.visit?.t_birds || 0,
          wks,
          r.b_weight_g,
          std?.weight || "-",
          r.feed_intake_g,
          std?.feed_bird_day || "-",
          r.egg_production,
          std?.egg_production || "-",
          r.egg_weight,
          std?.egg_weight || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          r.remarks || "-",
        ]);
      });
    } else if (reportType === "before_sale") {
      sheet.addRow([
        "Visit Date",
        "Agent",
        "Farmer / Batch",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Feed Type",
        "Remarks",
      ]);

      filteredReports.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = broilerStds[d];
        sheet.addRow([
          r.visit_date,
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.visit?.hatch_date || "-",
          r.visit?.t_birds || 0,
          d,
          r.mortality_pcs,
          r.weight_gm_b,
          std?.weight || "-",
          r.feed_t_kg,
          r.feed_gm_b,
          std?.feed_intake || "-",
          r.fcr_w_out_m,
          std?.fcr || "-",
          r.fcr_with_m,
          r.feed_company || "-",
          r.feed_type || "-",
          r.remarks || "-",
        ]);
      });
    } else {
      sheet.addRow([
        "Sale Date",
        "Agent",
        "Farmer / Batch",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt",
        "Total Feed (Kg)",
        "Feed/B (g)",
        "Std Feed",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Mill",
        "Feed Type",
        "Remarks",
      ]);

      filteredReports.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = broilerStds[d];
        const tB = Number(r.visit?.t_birds || 0);
        const fGm =
          r.feed_gm_b ||
          (tB > 0
            ? ((Number(r.feed_consumption_kg || 0) * 1000) / tB).toFixed(1)
            : "-");
        sheet.addRow([
          r.visit_date,
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.visit?.hatch_date || "-",
          tB,
          d,
          r.mortality,
          r.weight,
          std?.weight || "-",
          r.feed_consumption_kg,
          fGm,
          std?.feed_intake || "-",
          r.fcr_w_out_m,
          std?.fcr || "-",
          r.fcr_with_m,
          r.feed_mill || "-",
          r.feed || "-",
          r.remarks || "-",
        ]);
      });
    }

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buf]),
      `${selectedCategory}_${
        selectedCategory === "Layer" ? "Performance" : reportType
      }_Reports.xlsx`
    );
  }

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-4">
      {/* Header + Category & Report Type Tabs */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-white border border-slate-300 rounded-xl px-4 py-3 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          {(["Broiler", "Layer", "Sonali"] as const).map((cat) => (
            <button
              key={cat}
              type="button"
              onClick={() => setSelectedCategory(cat)}
              className={`px-3.5 py-1.5 rounded-lg text-xs font-bold transition ${
                selectedCategory === cat
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "bg-slate-100 text-slate-700 hover:bg-slate-200"
              }`}
            >
              {cat}
            </button>
          ))}

          {selectedCategory !== "Layer" ? (
            <div className="flex items-center gap-1.5 ml-2 pl-2 border-l border-slate-300">
              <button
                type="button"
                onClick={() => setReportType("before_sale")}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "before_sale"
                    ? "bg-slate-800 text-white"
                    : "bg-slate-100 text-slate-600"
                }`}
              >
                Before Sale
              </button>
              <button
                type="button"
                onClick={() => setReportType("after_sale")}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "after_sale"
                    ? "bg-red-600 text-white"
                    : "bg-slate-100 text-slate-600"
                }`}
              >
                After Sale
              </button>
            </div>
          ) : (
            <span className="ml-2 px-3 py-1 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs font-bold">
              Layer Performance Reports
            </span>
          )}
        </div>

        <button
          type="button"
          onClick={exportFilteredReportsExcel}
          className="bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm"
        >
          Export Excel ({filteredReports.length})
        </button>
      </div>

      {/* Filter Bar: Agent, Farmer, Month, From Date, To Date */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 items-center text-xs">
          <select
            value={agentFilter}
            onChange={(e) => {
              setAgentFilter(e.target.value);
              setFarmerFilter("all");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none"
          >
            <option value="all">All Agents</option>
            {agents.map((a) => (
              <option key={a.id} value={String(a.id)}>
                {a.agent_name}
              </option>
            ))}
          </select>

          <select
            value={farmerFilter}
            onChange={(e) => setFarmerFilter(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none font-semibold"
          >
            <option value="all">All Farmers</option>
            {availableFarmers.map((f) => (
              <option key={f.id} value={String(f.id)}>
                {f.farmer_name} ({f.village})
              </option>
            ))}
          </select>

          <select
            value={monthFilter}
            onChange={(e) => handleMonthSelect(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none"
          >
            <option value="">All Months</option>
            {allMonths.map((m) => (
              <option key={m} value={m}>
                Month: {m}
              </option>
            ))}
          </select>

          <div className="flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white">
            <span className="text-[10px] font-bold text-slate-400 mr-1">
              From:
            </span>
            <input
              type="date"
              value={fromDate}
              onChange={(e) => {
                setFromDate(e.target.value);
                setMonthFilter("");
              }}
              className="w-full outline-none bg-transparent"
            />
          </div>

          <div className="flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white">
            <span className="text-[10px] font-bold text-slate-400 mr-1">
              To:
            </span>
            <input
              type="date"
              value={toDate}
              onChange={(e) => {
                setToDate(e.target.value);
                setMonthFilter("");
              }}
              className="w-full outline-none bg-transparent"
            />
          </div>

          <button
            type="button"
            onClick={() => {
              setAgentFilter("all");
              setFarmerFilter("all");
              setMonthFilter("");
              setFromDate("");
              setToDate("");
            }}
            className="rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-3 py-1.5 font-semibold text-slate-700"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Batch-wise Grouped Tables */}
      {loading ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-500">
          Loading reports...
        </div>
      ) : groupedBatches.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-400">
          No reports found for the selected filters.
        </div>
      ) : (
        <div className="space-y-4">
          {groupedBatches.map((batch: any) => (
            <div
              key={batch.batchKey}
              className="bg-white border border-slate-300 rounded-xl overflow-hidden shadow-sm"
            >
              <div className="bg-slate-100 border-b border-slate-300 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
                <div>
                  <span className="font-bold text-slate-800 text-sm">
                    {batch.batchTitle}
                  </span>{" "}
                  <span className="text-slate-500">
                    • Agent: <b>{batch.agentName}</b>
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <span className="bg-white border border-slate-300 px-2 py-0.5 rounded">
                    H.Date: <b>{batch.hatchDate}</b>
                  </span>
                  <span className="bg-white border border-slate-300 px-2 py-0.5 rounded">
                    Birds: <b>{batch.tBirds}</b>
                  </span>
                  <span className="bg-white border border-slate-300 px-2 py-0.5 rounded">
                    Breed: <b>{batch.breed}</b>
                  </span>
                  <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded font-bold">
                    {batch.status}
                  </span>
                </div>
              </div>

              <div className="overflow-x-auto">
                {selectedCategory === "Layer" ? (
                  <table className="w-full border-collapse text-xs">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Age (Wks)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed / Std (g/b/d)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg% / Std%
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Egg Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Disease / Vaccine
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {batch.reports.map((r: any, i: number) => {
                        const wks = Math.max(
                          1,
                          Math.floor(
                            calcAgeDays(batch.hatchDate, r.visit_date) / 7
                          )
                        );
                        const std = layerStds[wks];
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="border border-slate-300 px-2 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                              {wks}w
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.b_weight_g}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.feed_intake_g}</b> /{" "}
                              {std?.feed_bird_day || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold text-indigo-700">
                              {r.egg_production}% / {std?.egg_production || "-"}
                              %
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.egg_weight}</b> / {std?.egg_weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.feed_company || "-"}{" "}
                              {r.feed_type ? `(${r.feed_type})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {[r.disease, r.vaccine]
                                .filter(Boolean)
                                .join(" / ") || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.remarks || "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : reportType === "before_sale" ? (
                  <table className="w-full border-collapse text-xs">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Age
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Mort.
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed (Kg)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed/B / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          FCR(w/o) / Std
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          FCR(With)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {batch.reports.map((r: any, i: number) => {
                        const d = calcAgeDays(batch.hatchDate, r.visit_date);
                        const std = broilerStds[d];
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="border border-slate-300 px-2 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                              {d}d
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.mortality_pcs}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.weight_gm_b}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.feed_t_kg}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.feed_gm_b}</b> / {std?.feed_intake || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold text-emerald-700">
                              {r.fcr_w_out_m} / {std?.fcr || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.fcr_with_m}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.feed_company || "-"}{" "}
                              {r.feed_type ? `(${r.feed_type})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.remarks || "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full border-collapse text-xs">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Sale Date
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Age
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Mort.
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Total Feed (Kg)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed/B / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          FCR(w/o) / Std
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          FCR(With)
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Feed Mill & Type
                        </th>
                        <th className="border border-slate-300 px-2 py-1.5">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {batch.reports.map((r: any, i: number) => {
                        const d = calcAgeDays(batch.hatchDate, r.visit_date);
                        const std = broilerStds[d];
                        const tB = Number(batch.tBirds || 0);
                        const fGm =
                          r.feed_gm_b ||
                          (tB > 0
                            ? (
                                (Number(r.feed_consumption_kg || 0) * 1000) /
                                tB
                              ).toFixed(1)
                            : "-");
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="border border-slate-300 px-2 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold">
                              {d}d
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.mortality}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{r.weight}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.feed_consumption_kg}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              <b>{fGm}</b> / {std?.feed_intake || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center font-bold text-blue-700">
                              {r.fcr_w_out_m} / {std?.fcr || "-"}
                            </td>
                            <td className="border border-slate-300 px-2 py-1 text-center">
                              {r.fcr_with_m}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.feed_mill || "-"} {r.feed ? `(${r.feed})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2 py-1">
                              {r.remarks || "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}