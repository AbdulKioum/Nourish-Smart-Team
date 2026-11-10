import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import ExcelJS from "exceljs";

import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";

function calcAgeDays(hatchDate?: string, visitDate?: string): number {
  if (!hatchDate) return 0;
  const hd = new Date(hatchDate);
  const vd = visitDate ? new Date(visitDate) : new Date();
  return Math.max(0, Math.floor((vd.getTime() - hd.getTime()) / 86400000));
}

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(String).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch {
      return val
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

function getRemarkText(rep: any): string {
  const r = rep?.remarks ?? rep?.remark ?? "";
  return String(r || "").trim();
}

export default function ReportsPage() {
  const [selectedCategory, setSelectedCategory] = useState<
    "all" | "Broiler" | "Layer" | "Sonali"
  >("all");
  const [reportType, setReportType] = useState<
    "all" | "before_sale" | "after_sale"
  >("all");

  const [userRole, setUserRole] = useState<string>("officer");
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);
  const [_assignedDistricts, setAssignedDistricts] = useState<string[]>([]);

  const [agents, setAgents] = useState<any[]>([]);
  const [farmers, setFarmers] = useState<any[]>([]);
  const [visitsMap, setVisitsMap] = useState<Record<string, any>>({});
  const [beforeReports, setBeforeReports] = useState<any[]>([]);
  const [afterReports, setAfterReports] = useState<any[]>([]);
  const [layerReports, setLayerReports] = useState<any[]>([]);
  const [broilerStds, setBroilerStds] = useState<Record<number, any>>({});
  const [sonaliStds, setSonaliStds] = useState<Record<number, any>>({});
  const [layerStds, setLayerStds] = useState<Record<string, any>>({});
  const [loading, setLoading] = useState(true);

  // Filters
  const [upazilaFilter, setUpazilaFilter] = useState<string>("all");
  const [agentFilter, setAgentFilter] = useState<string>("all");
  const [farmerFilter, setFarmerFilter] = useState<string>("all");
  const [monthFilter, setMonthFilter] = useState<string>("");
  const [isMultiDate, setIsMultiDate] = useState<boolean>(false);
  const [fromDate, setFromDate] = useState<string>("");
  const [toDate, setToDate] = useState<string>("");

  const isFullAccess = userRole === "super_admin" || userRole === "manager";

  useEffect(() => {
    loadReportsData();
  }, []);

  async function loadReportsData() {
    setLoading(true);

    const {
      data: { user },
    } = await supabase.auth.getUser();

    let role = "officer";
    let uUpazilas: string[] = [];
    let uDistricts: string[] = [];

    if (user) {
      const { data: uProf } = await supabase
        .from("users")
        .select("role, upazilas, districts")
        .eq("id", user.id)
        .single();

      if (uProf) {
        role = uProf.role || "officer";
        uUpazilas = parseList(uProf.upazilas);
        uDistricts = parseList(uProf.districts);
      }
    }

    setUserRole(role);
    setAssignedUpazilas(uUpazilas);
    setAssignedDistricts(uDistricts);

    const fullAccess = role === "super_admin" || role === "manager";
    if (!fullAccess && uUpazilas.length === 1) {
      setUpazilaFilter(uUpazilas[0]);
    }

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

    try {
      const { data: sStd } = await supabase
        .from("sonali_standards")
        .select("*");
      const sMap: Record<number, any> = {};
      (sStd || []).forEach((s: any) => {
        if (s.age_days !== undefined) sMap[Number(s.age_days)] = s;
      });
      setSonaliStds(sMap);
    } catch {
      setSonaliStds({});
    }

    const allFarmers = frmData || [];
    const allowedFarmers = fullAccess
      ? allFarmers
      : allFarmers.filter((f: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(f.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(f.district);
          return matchUpa && matchDist;
        });

    const allAgents = agData || [];
    const allowedAgents = fullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const aUpa = a.upazila || a.upazila_name || "";
          const matchUpa =
            uUpazilas.length === 0 || !aUpa || uUpazilas.includes(aUpa);
          const hasAllowedFarmer = allowedFarmers.some(
            (f: any) => String(f.agent_id) === String(a.id)
          );
          return matchUpa || hasAllowedFarmer;
        });

    setAgents(allowedAgents);
    setFarmers(allowedFarmers);

    const vMap: Record<string, any> = {};
    (vData || []).forEach((v: any) => (vMap[String(v.id)] = v));
    setVisitsMap(vMap);

    setBeforeReports(bRep || []);
    setAfterReports(aRep || []);
    setLayerReports(lRep || []);

    const bMap: Record<number, any> = {};
    (bStd || []).forEach((s: any) => (bMap[Number(s.age_days)] = s));
    setBroilerStds(bMap);

    const lMap: Record<string, any> = {};
    (lStd || []).forEach((s: any) => {
      const w = Number(s.age_weeks);
      const breedRaw = String(s.breed || "shaver star cross 579 brown")
        .trim()
        .toLowerCase();
      lMap[`${breedRaw}_${w}`] = s;
      if (breedRaw.includes("white")) {
        lMap[`white_${w}`] = s;
      } else {
        lMap[`brown_${w}`] = s;
      }
      if (!lMap[`any_${w}`]) {
        lMap[`any_${w}`] = s;
      }
    });
    setLayerStds(lMap);

    setLoading(false);
  }

  function getMeatStd(catName?: string, ageDays?: number) {
    const d = Number(ageDays || 0);
    const c = String(catName || "Broiler")
      .trim()
      .toLowerCase();
    if (c === "sonali") return sonaliStds[d] || null;
    return broilerStds[d] || null;
  }

  function getLayerStd(breedName?: string, ageWks?: number) {
    const w = Math.min(100, Math.max(1, Number(ageWks || 1)));
    const bLower = String(breedName || "")
      .trim()
      .toLowerCase();
    if (layerStds[`${bLower}_${w}`]) return layerStds[`${bLower}_${w}`];
    const isWhite =
      bLower.includes("white") ||
      bLower.includes("decalb") ||
      bLower.includes("bb 300");
    if (isWhite) {
      return (
        layerStds[`white_${w}`] ||
        layerStds[`bovans white_${w}`] ||
        layerStds[`any_${w}`] ||
        null
      );
    }
    return (
      layerStds[`brown_${w}`] ||
      layerStds[`shaver star cross 579 brown_${w}`] ||
      layerStds[`any_${w}`] ||
      null
    );
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
    setIsMultiDate(true);
    setFromDate(`${ym}-01`);
    setToDate(`${ym}-${String(lastDay).padStart(2, "0")}`);
  }

  // Upazila options strictly scoped to user's assigned upazilas
  const upazilaOptions = Array.from(
    new Set(
      (!isFullAccess && assignedUpazilas.length > 0
        ? assignedUpazilas
        : [
            ...farmers.map((f) => f.upazila),
            ...agents.map((a) => a.upazila || a.upazila_name),
          ]
      ).filter(Boolean)
    )
  ).sort();

  const availableAgents = agents.filter((a) => {
    if (upazilaFilter === "all") return true;
    const aUpa = a.upazila || a.upazila_name || "";
    const hasFarmerInUpa = farmers.some(
      (f) =>
        String(f.agent_id) === String(a.id) && f.upazila === upazilaFilter
    );
    return aUpa === upazilaFilter || hasFarmerInUpa;
  });

  const availableFarmers = farmers.filter((f) => {
    if (upazilaFilter !== "all" && f.upazila !== upazilaFilter) return false;
    if (agentFilter !== "all" && String(f.agent_id) !== String(agentFilter))
      return false;
    return true;
  });

  // Combine raw reports based on selectedCategory and reportType
  const rawCombinedReports: any[] = [];
  if (selectedCategory === "all") {
    if (reportType === "all" || reportType === "before_sale") {
      beforeReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "before_sale" })
      );
    }
    if (reportType === "all" || reportType === "after_sale") {
      afterReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "after_sale" })
      );
    }
    if (reportType === "all") {
      layerReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "layer" })
      );
    }
  } else if (selectedCategory === "Layer") {
    layerReports.forEach((r) =>
      rawCombinedReports.push({ ...r, _repKind: "layer" })
    );
  } else {
    if (reportType === "all" || reportType === "before_sale") {
      beforeReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "before_sale" })
      );
    }
    if (reportType === "all" || reportType === "after_sale") {
      afterReports.forEach((r) =>
        rawCombinedReports.push({ ...r, _repKind: "after_sale" })
      );
    }
  }

  const allowedFarmerIds = new Set(farmers.map((f) => String(f.id)));

  // Enrich and Filter Reports
  const filteredReports = rawCombinedReports
    .map((r) => {
      const visit = visitsMap[String(r.vet_visit_id)] || {};
      const farmer =
        farmers.find((f) => String(f.id) === String(r.farmer_id)) || {};
      const agent =
        agents.find(
          (a) => String(a.id) === String(r.agent_id || farmer.agent_id)
        ) || {};
      const cat =
        r._repKind === "layer"
          ? "Layer"
          : visit.category || farmer.farm_type || "Broiler";
      const upazila =
        farmer.upazila || agent.upazila || agent.upazila_name || "";
      const hatchery = r.hatchery || visit.hatchery || "-";
      const breed = r.breed || visit.breed || "-";

      return {
        ...r,
        visit,
        farmer,
        agent,
        upazila,
        hatchery,
        breed,
        category: cat,
        batchKey: `${r._repKind}_${String(r.vet_visit_id || r.farmer_id)}`,
        batchTitle:
          visit.batch_label || farmer.farmer_name || "Unknown Farmer",
      };
    })
    .filter((item) => {
      if (!isFullAccess && assignedUpazilas.length > 0) {
        const matchAllowed =
          allowedFarmerIds.has(String(item.farmer_id)) ||
          assignedUpazilas.includes(item.upazila);
        if (!matchAllowed) return false;
      }
      if (
        selectedCategory !== "all" &&
        item.category?.toLowerCase() !== selectedCategory.toLowerCase()
      ) {
        return false;
      }
      if (upazilaFilter !== "all" && item.upazila !== upazilaFilter) {
        return false;
      }
      if (
        agentFilter !== "all" &&
        String(item.agent_id || item.farmer?.agent_id) !== String(agentFilter)
      ) {
        return false;
      }
      if (
        farmerFilter !== "all" &&
        String(item.farmer_id) !== String(farmerFilter)
      ) {
        return false;
      }
      const vDate = String(item.visit_date || "").slice(0, 10);
      if (fromDate && vDate < fromDate) return false;
      if (toDate && vDate > toDate) return false;
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

  // Group by Batch + Report Kind
  const groupedBatches = Object.values(
    filteredReports.reduce((acc: Record<string, any>, rep: any) => {
      const key = rep.batchKey;
      if (!acc[key]) {
        acc[key] = {
          batchKey: key,
          repKind: rep._repKind,
          category: rep.category,
          batchTitle: rep.batchTitle,
          agentName: rep.agent?.agent_name || "-",
          upazila: rep.upazila || "-",
          farmerName: rep.farmer?.farmer_name || "-",
          hatchery: rep.hatchery || "-",
          hatchDate: rep.visit?.hatch_date || "-",
          breed: rep.breed || "-",
          tBirds: rep.visit?.t_birds || 0,
          status: rep.visit?.status || "-",
          reports: [],
        };
      }
      acc[key].reports.push(rep);
      return acc;
    }, {})
  );

  // Professional Excel Worksheet Formatter
  function formatExcelSheet(
    sheet: ExcelJS.Worksheet,
    headerHex: string,
    numericCols: number[]
  ) {
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    const headerRow = sheet.getRow(1);
    headerRow.height = 26;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerHex },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: false,
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFFFFFFF" } },
        left: { style: "thin", color: { argb: "FFFFFFFF" } },
        bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
        right: { style: "thin", color: { argb: "FFFFFFFF" } },
      };
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      row.height = 20;
      const isEven = rowNumber % 2 === 0;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (isEven) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF8FAFC" },
          };
        }
        cell.font = { size: 10, color: { argb: "FF0F172A" } };
        cell.alignment = {
          vertical: "middle",
          horizontal: numericCols.includes(colNumber) ? "right" : "left",
          wrapText: false,
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
      });
    });

    sheet.columns.forEach((column) => {
      let maxLen = 11;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        const val = cell.value ? String(cell.value) : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      column.width = Math.min(Math.max(maxLen + 4, 12), 38);
    });
  }


async function saveBufferOrBlob(
  bufferOrBlob: ArrayBuffer | Blob,
  fileName: string
) {
  try {
    const arrayBuffer =
      bufferOrBlob instanceof Blob
        ? await bufferOrBlob.arrayBuffer()
        : bufferOrBlob;

    if (Capacitor.isNativePlatform()) {
      let binary = "";
      const bytes = new Uint8Array(arrayBuffer);
      const chunkSize = 0x8000;
      for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode.apply(
          null,
          Array.from(bytes.subarray(i, i + chunkSize))
        );
      }
      const base64Data = window.btoa(binary);

      const savedFile = await Filesystem.writeFile({
        path: fileName,
        data: base64Data,
        directory: Directory.Documents,
        recursive: true,
      });

      alert(`File successfully saved:\n📁 Documents/${fileName}`);

      if (savedFile.uri) {
        try {
          await Share.share({
            title: fileName,
            url: savedFile.uri,
            dialogTitle: "Open or View Excel File",
          });
        } catch {}
      }
    } else {
      const blob =
        bufferOrBlob instanceof Blob
          ? bufferOrBlob
          : new Blob([arrayBuffer], {
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }
  } catch (err: any) {
    console.error("Export save error:", err);
    alert("Failed to save file: " + (err?.message || ""));
  }
}

  async function exportFilteredReportsExcel() {
    if (filteredReports.length === 0) {
      alert("No reports found to export.");
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const beforeList = filteredReports.filter(
      (r) => r._repKind === "before_sale"
    );
    const afterList = filteredReports.filter(
      (r) => r._repKind === "after_sale"
    );
    const layerList = filteredReports.filter((r) => r._repKind === "layer");

    if (beforeList.length > 0) {
      const s = workbook.addWorksheet("Before_Sale_Reports");
      s.addRow([
        "Visit Date",
        "Category",
        "Upazila",
        "Agent",
        "Farmer / Batch",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Feed Type",
        "Remarks",
      ]);

      beforeList.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = getMeatStd(r.category, d);
        s.addRow([
          r.visit_date || "-",
          r.category || "-",
          r.upazila || "-",
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          Number(r.visit?.t_birds || 0),
          d,
          Number(r.mortality_pcs || 0),
          Number(r.weight_gm_b || 0),
          std?.weight || "-",
          Number(r.feed_t_kg || 0),
          Number(r.feed_gm_b || 0),
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FF059669",
        [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      );
    }

    if (afterList.length > 0) {
      const s = workbook.addWorksheet("After_Sale_Reports");
      s.addRow([
        "Sale Date",
        "Category",
        "Upazila",
        "Agent",
        "Farmer / Batch",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Age (Days)",
        "Mortality",
        "Weight (g)",
        "Std Wt (g)",
        "Total Feed (Kg)",
        "Feed/B (g)",
        "Std Feed (g)",
        "FCR (w/o M)",
        "Std FCR",
        "FCR (With M)",
        "Feed Company",
        "Last Feed",
        "Remarks",
      ]);

      afterList.forEach((r: any) => {
        const d = calcAgeDays(r.visit?.hatch_date, r.visit_date);
        const std = getMeatStd(r.category, d);
        const tB = Number(r.visit?.t_birds || 0);
        const fGm =
          r.feed_gm_b ||
          (tB > 0
            ? ((Number(r.feed_consumption_kg || 0) * 1000) / tB).toFixed(1)
            : "-");
        s.addRow([
          r.visit_date || "-",
          r.category || "-",
          r.upazila || "-",
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          tB,
          d,
          Number(r.mortality || 0),
          Number(r.weight || 0),
          std?.weight || "-",
          Number(r.feed_consumption_kg || 0),
          fGm,
          std?.feed_intake || "-",
          r.fcr_w_out_m || "-",
          std?.fcr || "-",
          r.fcr_with_m || "-",
          r.feed_mill || "-",
          r.feed || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FFDC2626",
        [9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      );
    }

    if (layerList.length > 0) {
      const s = workbook.addWorksheet("Layer_Reports");
      s.addRow([
        "Visit Date",
        "Upazila",
        "Agent",
        "Farmer / Batch",
        "Hatchery",
        "Breed",
        "H.Date",
        "T.Birds",
        "Mortality",
        "Age (Wks)",
        "Weight (g)",
        "Std Wt (g)",
        "Feed (g/b/d)",
        "Std Feed (g)",
        "Egg Prod (%)",
        "Std Egg (%)",
        "Egg Wt (g)",
        "Std Egg Wt (g)",
        "Feed Company",
        "Feed Type",
        "Disease / Vaccine",
        "Remarks",
      ]);

      layerList.forEach((r: any) => {
        const wks = Math.max(
          1,
          Math.floor(calcAgeDays(r.visit?.hatch_date, r.visit_date) / 7)
        );
        const std = getLayerStd(r.breed, wks);
        s.addRow([
          r.visit_date || "-",
          r.upazila || "-",
          r.agent?.agent_name || "-",
          r.batchTitle,
          r.hatchery || "-",
          r.breed || "-",
          r.visit?.hatch_date || "-",
          Number(r.visit?.t_birds || 0),
          Number(r.mortality ?? 0),
          wks,
          Number(r.b_weight_g || 0),
          std?.weight || "-",
          Number(r.feed_intake_g || 0),
          std?.feed_bird_day || "-",
          Number(r.egg_production || 0),
          std?.egg_production || "-",
          Number(r.egg_weight || 0),
          std?.egg_weight || "-",
          r.feed_company || "-",
          r.feed_type || "-",
          [r.disease, r.vaccine].filter(Boolean).join(" / ") || "-",
          getRemarkText(r) || "-",
        ]);
      });

      formatExcelSheet(
        s,
        "FF4F46E5",
        [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
      );
    }

    const today = new Date().toISOString().slice(0, 10);
    const fileName = `SmartTeam_Reports_${selectedCategory || "All"}_${today}.xlsx`;

    try {
      const excelBuffer = await workbook.xlsx.writeBuffer();
      // DOC Tracking বা Farmer পেজে তৈরি করা common saveBufferOrBlob ফাংশনটি এখানে কল করা হলো
      await saveBufferOrBlob(excelBuffer, fileName);
    } catch (err: any) {
      console.error("Filtered reports export error:", err);
      alert("Reports export failed.");
    }
  }

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-3.5">
      {/* Top Bar: All + Category Tabs & Report Type Tabs */}
      <div className="flex flex-wrap items-center justify-between gap-2.5 bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5">
          {(["all", "Broiler", "Layer", "Sonali"] as const).map((cat) => (
            <button
              key={cat}
              type="button"
              onClick={() => setSelectedCategory(cat)}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                selectedCategory === cat
                  ? "bg-emerald-600 text-white shadow-sm"
                  : "bg-slate-100 text-slate-700 hover:bg-slate-200"
              }`}
            >
              {cat === "all" ? "All" : cat}
            </button>
          ))}

          {selectedCategory !== "Layer" ? (
            <div className="flex items-center gap-1.5 ml-1.5 pl-2 border-l border-slate-300">
              <button
                type="button"
                onClick={() => setReportType("all")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "all"
                    ? "bg-indigo-600 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                All
              </button>
              <button
                type="button"
                onClick={() => setReportType("before_sale")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "before_sale"
                    ? "bg-slate-800 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                Before Sale
              </button>
              <button
                type="button"
                onClick={() => setReportType("after_sale")}
                className={`px-2.5 py-1.5 rounded-lg text-xs font-semibold ${
                  reportType === "after_sale"
                    ? "bg-red-600 text-white"
                    : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}
              >
                After Sale
              </button>
            </div>
          ) : (
            <span className="ml-1.5 px-2.5 py-1 rounded-lg bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs font-bold">
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

      {/* Filter Bar: Assigned Upazila, Agent, Farmer, Month, Single/Multi Date Picker, Reset */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div
          className={`grid grid-cols-2 sm:grid-cols-3 ${
            isMultiDate ? "lg:grid-cols-7" : "lg:grid-cols-6"
          } gap-2 items-center text-xs`}
        >
          {/* Assigned Upazila Dropdown */}
          <select
            value={upazilaFilter}
            onChange={(e) => {
              setUpazilaFilter(e.target.value);
              setAgentFilter("all");
              setFarmerFilter("all");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none"
          >
            <option value="all">All Upazilas</option>
            {upazilaOptions.map((u) => (
              <option key={u} value={u}>
                {u}
              </option>
            ))}
          </select>

          {/* Agent Dropdown */}
          <select
            value={agentFilter}
            onChange={(e) => {
              setAgentFilter(e.target.value);
              setFarmerFilter("all");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none"
          >
            <option value="all">All Agents ({availableAgents.length})</option>
            {availableAgents.map((a) => (
              <option key={a.id} value={String(a.id)}>
                {a.agent_name}
              </option>
            ))}
          </select>

          {/* Farmer Dropdown */}
          <select
            value={farmerFilter}
            onChange={(e) => setFarmerFilter(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 outline-none font-semibold"
          >
            <option value="all">All Farmers ({availableFarmers.length})</option>
            {availableFarmers.map((f) => (
              <option key={f.id} value={String(f.id)}>
                {f.farmer_name} ({f.village})
              </option>
            ))}
          </select>

          {/* Month Filter */}
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

          {/* Single Date OR Multi (From - To) Date Selector */}
          <div
            className={`${
              isMultiDate
                ? "col-span-2 sm:col-span-3 lg:col-span-2"
                : "col-span-2 sm:col-span-2 lg:col-span-1"
            } flex items-center gap-1.5 min-w-0`}
          >
            {!isMultiDate ? (
              <div className="flex-1 min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                <span className="text-[10px] font-bold text-slate-400 mr-1.5 shrink-0">
                  Date:
                </span>
                <input
                  type="date"
                  value={fromDate}
                  onChange={(e) => {
                    const val = e.target.value;
                    setFromDate(val);
                    setToDate(val);
                    setMonthFilter("");
                  }}
                  className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                />
              </div>
            ) : (
              <div className="flex-1 min-w-0 grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    From:
                  </span>
                  <input
                    type="date"
                    value={fromDate}
                    onChange={(e) => {
                      setFromDate(e.target.value);
                      setMonthFilter("");
                    }}
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
                <div className="min-w-0 flex items-center border border-slate-300 rounded-lg px-2 py-1 bg-white overflow-hidden">
                  <span className="text-[10px] font-bold text-slate-400 mr-1 shrink-0">
                    To:
                  </span>
                  <input
                    type="date"
                    value={toDate}
                    onChange={(e) => {
                      setToDate(e.target.value);
                      setMonthFilter("");
                    }}
                    className="w-full min-w-0 outline-none bg-transparent text-slate-700"
                  />
                </div>
              </div>
            )}

            <label
              title="Select multiple dates (From - To)"
              className="shrink-0 inline-flex items-center gap-1 px-2 py-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-[10px] font-semibold text-slate-700 cursor-pointer select-none"
            >
              <input
                type="checkbox"
                checked={isMultiDate}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setIsMultiDate(checked);
                  if (!checked) {
                    const single = fromDate || toDate || "";
                    setFromDate(single);
                    setToDate(single);
                    setMonthFilter("");
                  }
                }}
                className="w-3.5 h-3.5 accent-emerald-600"
              />
              <span>Multi</span>
            </label>
          </div>

          {/* Reset Button */}
          <button
            type="button"
            onClick={() => {
              setUpazilaFilter(
                !isFullAccess && assignedUpazilas.length === 1
                  ? assignedUpazilas[0]
                  : "all"
              );
              setAgentFilter("all");
              setFarmerFilter("all");
              setMonthFilter("");
              setIsMultiDate(false);
              setFromDate("");
              setToDate("");
            }}
            className="rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-3 py-1.5 font-semibold text-slate-700"
          >
            Reset
          </button>
        </div>
      </div>

      {/* Batch-wise Grouped Tables (Single-line rows with Hatchery in batch info bar) */}
      {loading ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-500">
          Loading reports...
        </div>
      ) : groupedBatches.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-xl p-8 text-center text-xs text-slate-400">
          No reports found for the selected filters.
        </div>
      ) : (
        <div className="space-y-3.5">
          {groupedBatches.map((batch: any) => (
            <div
              key={batch.batchKey}
              className="bg-white border border-slate-300 rounded-xl overflow-hidden shadow-sm"
            >
              {/* Farmer / Batch Info Header (Includes Hatchery) */}
              <div className="bg-slate-100 border-b border-slate-300 px-3.5 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="font-bold text-slate-800 text-sm">
                    {batch.batchTitle}
                  </span>
                  <span className="px-1.5 py-0.5 rounded bg-slate-200 text-slate-700 text-[10px] font-bold">
                    {batch.category} •{" "}
                    {batch.repKind === "layer"
                      ? "Layer Report"
                      : batch.repKind === "after_sale"
                      ? "After Sale"
                      : "Before Sale"}
                  </span>
                  <span className="text-slate-500">
                    • Agent: <b>{batch.agentName}</b> ({batch.upazila})
                  </span>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="bg-white border border-slate-300 px-2 py-0.5 rounded">
                    Hatchery: <b className="text-emerald-800">{batch.hatchery}</b>
                  </span>
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
                    {batch.status === "Activated" ? "Running" : batch.status}
                  </span>
                </div>
              </div>

              <div className="overflow-x-auto">
                {batch.repKind === "layer" ? (
                  <table className="w-full border-collapse text-xs whitespace-nowrap">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Age (Wks)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Mort.
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Feed / Std (g/b/d)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Egg% / Std%
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Egg Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Disease / Vaccine
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
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
                        const std = getLayerStd(r.breed || batch.breed, wks);
                        return (
                          <tr key={i} className="hover:bg-slate-50 leading-tight">
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold">
                              {wks}w
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.mortality ?? 0}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.b_weight_g}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.feed_intake_g}</b> /{" "}
                              {std?.feed_bird_day || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold text-indigo-700">
                              {r.egg_production}% / {std?.egg_production || "-"}
                              %
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.egg_weight}</b> / {std?.egg_weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.feed_company || "-"}{" "}
                              {r.feed_type ? `(${r.feed_type})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {[r.disease, r.vaccine]
                                .filter(Boolean)
                                .join(" / ") || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {getRemarkText(r) || "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : batch.repKind === "before_sale" ? (
                  <table className="w-full border-collapse text-xs whitespace-nowrap">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Visit Date
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Age
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Mort.
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Feed (Kg)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Feed/B / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          FCR(w/o) / Std
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          FCR(With)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Feed & Type
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {batch.reports.map((r: any, i: number) => {
                        const d = calcAgeDays(batch.hatchDate, r.visit_date);
                        const std = getMeatStd(batch.category, d);
                        return (
                          <tr key={i} className="hover:bg-slate-50 leading-tight">
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold">
                              {d}d
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.mortality_pcs}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.weight_gm_b}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.feed_t_kg}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.feed_gm_b}</b> / {std?.feed_intake || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold text-emerald-700">
                              {r.fcr_w_out_m} / {std?.fcr || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.fcr_with_m}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.feed_company || "-"}{" "}
                              {r.feed_type ? `(${r.feed_type})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {getRemarkText(r) || "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full border-collapse text-xs whitespace-nowrap">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Sale Date
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Age
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Mort.
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Wt / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Total Feed (Kg)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          Feed/B / Std (g)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          FCR(w/o) / Std
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-center">
                          FCR(With)
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Feed Company & Last Feed
                        </th>
                        <th className="border border-slate-300 px-2.5 py-1.5 text-left">
                          Remarks
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {batch.reports.map((r: any, i: number) => {
                        const d = calcAgeDays(batch.hatchDate, r.visit_date);
                        const std = getMeatStd(batch.category, d);
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
                          <tr key={i} className="hover:bg-slate-50 leading-tight">
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.visit_date}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold">
                              {d}d
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.mortality}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{r.weight}</b> / {std?.weight || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.feed_consumption_kg}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              <b>{fGm}</b> / {std?.feed_intake || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center font-bold text-blue-700">
                              {r.fcr_w_out_m} / {std?.fcr || "-"}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1 text-center">
                              {r.fcr_with_m}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {r.feed_mill || "-"}{" "}
                              {r.feed ? `(Last Feed: ${r.feed})` : ""}
                            </td>
                            <td className="border border-slate-300 px-2.5 py-1">
                              {getRemarkText(r) || "-"}
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