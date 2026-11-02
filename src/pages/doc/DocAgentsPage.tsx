import React, { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import {
  FaUserTie,
  FaPlus,
  FaEdit,
  FaTrash,
  FaSearch,
  FaTimes,
  FaSpinner,
  FaLink,
} from "react-icons/fa";
import { RotateCcw } from "lucide-react";

interface DocAgent {
  id: string;
  doc_id: string;
  agent_name: string;
  proprietor: string;
  mobile: string;
  district: string;
  upazila: string;
  feed_agent_id?: number | null;
}

interface AgentTableRecord {
  id: number;
  agent_code?: string;
  agent_name: string;
  proprietor?: string;
  mobile?: string;
  upazila?: string;
  district?: string;
  companies?: { company_name?: string };
}

interface DistrictItem {
  id: string | number;
  name: string;
}

interface UpazilaItem {
  id: string | number;
  name: string;
  district_id?: string | number;
  district_name?: string;
  district?: string;
}

export default function DocAgentsPage() {
  const [agents, setAgents] = useState<DocAgent[]>([]);
  const [feedAgents, setFeedAgents] = useState<AgentTableRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");
  const [successMsg, setSuccessMsg] = useState("");

  const [userRole, setUserRole] = useState("officer");
  const [assignedDistricts, setAssignedDistricts] = useState<string[]>([]);
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);

  const [districtsList, setDistrictsList] = useState<DistrictItem[]>([]);
  const [upazilasList, setUpazilasList] = useState<UpazilaItem[]>([]);

  const [search, setSearch] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingAgent, setEditingAgent] = useState<DocAgent | null>(null);

  const [formData, setFormData] = useState({
    doc_id: "",
    agent_name: "",
    proprietor: "",
    mobile: "",
    district: "",
    upazila: "",
    feed_agent_id: "",
  });

  const canManage = userRole === "super_admin" || userRole === "manager";

  useEffect(() => {
    initPage();
  }, []);

  function parseArray(field: any): string[] {
    if (!field) return [];
    if (Array.isArray(field)) return field;
    if (typeof field === "string") {
      try {
        const parsed = JSON.parse(field);
        if (Array.isArray(parsed)) return parsed;
      } catch {
        return field
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);
      }
    }
    return [];
  }

  async function initPage() {
    setLoading(true);
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();

      let role = "officer";
      let uDistricts: string[] = [];
      let uUpazilas: string[] = [];

      if (user) {
        const { data: profile } = await supabase
          .from("users")
          .select("role, districts, upazilas")
          .eq("id", user.id)
          .single();

        if (profile) {
          role = profile.role || "officer";
          uDistricts = parseArray(profile.districts);
          uUpazilas = parseArray(profile.upazilas);
        }
      }

      setUserRole(role);
      setAssignedDistricts(uDistricts);
      setAssignedUpazilas(uUpazilas);

      const [{ data: dData }, { data: uData }, { data: faData }] =
        await Promise.all([
          supabase.from("districts").select("*").order("name"),
          supabase.from("upazilas").select("*").order("name"),
          supabase
            .from("agents")
            .select(
              "id, agent_code, agent_name, proprietor, mobile, upazila, district, companies(company_name)"
            )
            .order("agent_name"),
        ]);

      setDistrictsList(dData || []);
      setUpazilasList(uData || []);
      setFeedAgents((faData as any[]) || []);

      await fetchAgents(role, uUpazilas);
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to load data.");
    } finally {
      setLoading(false);
    }
  }

  async function fetchAgents(role = userRole, uUpazilas = assignedUpazilas) {
    let query = supabase
      .from("doc_agents")
      .select("*")
      .order("doc_id", { ascending: true });

    if (role !== "super_admin" && role !== "manager") {
      if (uUpazilas.length > 0) {
        query = query.in("upazila", uUpazilas);
      } else {
        setAgents([]);
        return;
      }
    }

    const { data, error } = await query;
    if (error) throw error;
    setAgents(data || []);
  }

  const allowedDistricts = districtsList.filter((d) => {
    if (canManage || assignedDistricts.length === 0) return true;
    return assignedDistricts.includes(d.name);
  });

  const allowedUpazilasForSelectedDistrict = upazilasList.filter((u) => {
    if (!formData.district) return false;

    const matchedDist = districtsList.find(
      (d) => String(d.id) === String(u.district_id)
    );
    const dName = matchedDist?.name || u.district_name || u.district;

    if (dName !== formData.district) return false;
    if (canManage) return true;
    return assignedUpazilas.includes(u.name);
  });

  // Strictly show agents of the selected Upazila from `agents` table
  const upazilaAgentsForCollaboration = feedAgents.filter((fa) => {
    if (formData.upazila) {
      return (
        String(fa.upazila || "").trim().toLowerCase() ===
        formData.upazila.trim().toLowerCase()
      );
    }
    if (formData.district) {
      return (
        String(fa.district || "").trim().toLowerCase() ===
        formData.district.trim().toLowerCase()
      );
    }
    return true;
  });

  function handleOpenCreate() {
    if (!canManage) return;
    setEditingAgent(null);
    setErrorMsg("");
    setSuccessMsg("");

    const defaultDistrict =
      allowedDistricts.length === 1 ? allowedDistricts[0].name : "Mymensingh";
    const defaultUpazila =
      assignedUpazilas.length === 1 ? assignedUpazilas[0] : "";

    setFormData({
      doc_id: "",
      agent_name: "",
      proprietor: "",
      mobile: "",
      district: defaultDistrict,
      upazila: defaultUpazila,
      feed_agent_id: "",
    });
    setIsModalOpen(true);
  }

  function handleOpenEdit(agent: DocAgent) {
    if (!canManage) return;
    setEditingAgent(agent);
    setErrorMsg("");
    setSuccessMsg("");
    setFormData({
      doc_id: agent.doc_id,
      agent_name: agent.agent_name,
      proprietor: agent.proprietor || "",
      mobile: agent.mobile || "",
      district: agent.district || "",
      upazila: agent.upazila || "",
      feed_agent_id: agent.feed_agent_id ? String(agent.feed_agent_id) : "",
    });
    setIsModalOpen(true);
  }

  function handleSelectCollaboratedAgent(selectedId: string) {
    const found = feedAgents.find((fa) => String(fa.id) === selectedId);
    if (!found) {
      setFormData((prev) => ({ ...prev, feed_agent_id: "" }));
      return;
    }

    setFormData((prev) => ({
      ...prev,
      feed_agent_id: selectedId,
      doc_id: prev.doc_id || found.agent_code || "",
      agent_name: prev.agent_name || found.agent_name || "",
      proprietor: prev.proprietor || found.proprietor || "",
      mobile: prev.mobile || found.mobile || "",
    }));
  }

  async function handleSave(e: React.FormEvent) {
    e.preventDefault();
    if (!canManage) return;
    setSubmitting(true);
    setErrorMsg("");

    try {
      const cleanDocId = formData.doc_id.trim();
      if (!cleanDocId || !formData.agent_name.trim()) {
        throw new Error("DOC.ID and Agent Name are required.");
      }
      if (!formData.district || !formData.upazila) {
        throw new Error("District and Upazila are required.");
      }

      const payload: Record<string, any> = {
        doc_id: cleanDocId,
        agent_name: formData.agent_name.trim(),
        proprietor: formData.proprietor.trim(),
        mobile: formData.mobile.trim(),
        district: formData.district,
        upazila: formData.upazila,
        feed_agent_id: formData.feed_agent_id
          ? Number(formData.feed_agent_id)
          : null,
      };

      if (editingAgent) {
        const { error } = await supabase
          .from("doc_agents")
          .update(payload)
          .eq("id", editingAgent.id);

        if (error) {
          delete payload.feed_agent_id;
          const { error: retryErr } = await supabase
            .from("doc_agents")
            .update(payload)
            .eq("id", editingAgent.id);
          if (retryErr) throw retryErr;
        }
        setSuccessMsg("DOC Agent updated successfully.");
      } else {
        const { error } = await supabase.from("doc_agents").insert([payload]);

        if (error) {
          if (error.code === "23505") {
            throw new Error(`DOC.ID "${cleanDocId}" already exists!`);
          }
          delete payload.feed_agent_id;
          const { error: retryErr } = await supabase
            .from("doc_agents")
            .insert([payload]);
          if (retryErr) throw retryErr;
        }
        setSuccessMsg("DOC Agent added successfully.");
      }

      await fetchAgents();
      setIsModalOpen(false);
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to save DOC Agent.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDelete(agent: DocAgent) {
    if (!canManage) return;
    if (
      !window.confirm(
        `Delete DOC Agent "${agent.agent_name}" (ID: ${agent.doc_id})?`
      )
    )
      return;

    try {
      const { error } = await supabase
        .from("doc_agents")
        .delete()
        .eq("id", agent.id);
      if (error) throw error;
      await fetchAgents();
    } catch (err: any) {
      setErrorMsg(err?.message || "Failed to delete agent.");
    }
  }

  function resetFilters() {
    setSearch("");
    setUpazilaFilter("");
  }

  const filteredAgents = agents.filter((a) => {
    const q = search.trim().toLowerCase();
    const matchSearch =
      !q ||
      a.doc_id.toLowerCase().includes(q) ||
      a.agent_name.toLowerCase().includes(q) ||
      a.proprietor?.toLowerCase().includes(q) ||
      a.mobile?.toLowerCase().includes(q);

    const matchUpazila = !upazilaFilter || a.upazila === upazilaFilter;
    return matchSearch && matchUpazila;
  });

  const uniqueUpazilas = Array.from(new Set(agents.map((a) => a.upazila)));

  function getCollaboratedAgentName(agent: DocAgent) {
    if (agent.feed_agent_id) {
      const found = feedAgents.find(
        (fa) => String(fa.id) === String(agent.feed_agent_id)
      );
      if (found) {
        return `${found.agent_name}${
          found.companies?.company_name
            ? ` (${found.companies.company_name})`
            : ""
        }`;
      }
    }
    const byCodeOrName = feedAgents.find(
      (fa) =>
        (fa.agent_code && fa.agent_code === agent.doc_id) ||
        fa.agent_name.trim().toLowerCase() ===
          agent.agent_name.trim().toLowerCase()
    );
    return byCodeOrName ? `${byCodeOrName.agent_name} (Auto)` : "-";
  }

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-200 rounded-xl px-3.5 py-2.5 shadow-sm">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-emerald-100 text-emerald-600 flex items-center justify-center shrink-0">
            <FaUserTie className="text-sm" />
          </div>
          <h1 className="text-base sm:text-lg font-bold text-slate-800">
            DOC Agents
          </h1>
          <span className="px-2.5 py-0.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-semibold">
            Total: {filteredAgents.length}
          </span>
        </div>

        {canManage && (
          <button
            onClick={handleOpenCreate}
            className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <FaPlus size={11} />
            Add DOC Agent
          </button>
        )}
      </div>

      {errorMsg && (
        <div className="flex items-center justify-between bg-red-50 border border-red-200 text-red-700 px-3 py-2 rounded-xl text-xs">
          <span>{errorMsg}</span>
          <button onClick={() => setErrorMsg("")}>
            <FaTimes />
          </button>
        </div>
      )}

      {successMsg && (
        <div className="flex items-center justify-between bg-emerald-50 border border-emerald-200 text-emerald-700 px-3 py-2 rounded-xl text-xs">
          <span>{successMsg}</span>
          <button onClick={() => setSuccessMsg("")}>
            <FaTimes />
          </button>
        </div>
      )}

      {/* Compact Search & Filter Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-2.5 shadow-sm">
        <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
          <div className="relative flex-1 min-w-[180px]">
            <FaSearch className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-xs" />
            <input
              type="text"
              placeholder="Search DOC.ID, Agent Name, Proprietor or Mobile..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full border border-slate-300 rounded-lg pl-8 pr-3 py-1.5 text-xs outline-none focus:border-emerald-500"
            />
          </div>

          <select
            value={upazilaFilter}
            onChange={(e) => setUpazilaFilter(e.target.value)}
            className="border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs outline-none bg-white"
          >
            <option value="">All Upazilas</option>
            {uniqueUpazilas.map((u) => (
              <option key={u} value={u}>
                {u}
              </option>
            ))}
          </select>

          <button
            onClick={resetFilters}
            title="Reset Filters"
            className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2.5 py-1.5 text-xs font-semibold text-slate-700 transition"
          >
            <RotateCcw size={13} />
          </button>
        </div>
      </div>

      {/* Simple Bordered Table */}
      <div className="bg-white border border-slate-300 rounded-xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="min-w-full border-collapse text-xs">
            <thead className="bg-slate-100 border-b border-slate-300 text-slate-700">
              <tr>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  DOC.ID
                </th>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  Agent Name
                </th>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  Collaborated Agent (Agents Page)
                </th>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  Proprietor
                </th>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  Mobile
                </th>
                <th className="border border-slate-300 px-3 py-2 text-left">
                  Upazila & District
                </th>
                {canManage && (
                  <th className="border border-slate-300 px-3 py-2 text-center">
                    Actions
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td
                    colSpan={canManage ? 7 : 6}
                    className="py-8 text-center text-slate-500"
                  >
                    <FaSpinner className="animate-spin inline mr-2" /> Loading...
                  </td>
                </tr>
              ) : filteredAgents.length === 0 ? (
                <tr>
                  <td
                    colSpan={canManage ? 7 : 6}
                    className="py-8 text-center text-slate-400"
                  >
                    No DOC Agents found.
                  </td>
                </tr>
              ) : (
                filteredAgents.map((agent) => (
                  <tr key={agent.id} className="hover:bg-slate-50">
                    <td className="border border-slate-300 px-3 py-1.5 font-bold text-emerald-700">
                      {agent.doc_id}
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold text-slate-800">
                      {agent.agent_name}
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 text-indigo-700 font-medium">
                      <span className="inline-flex items-center gap-1">
                        <FaLink size={10} />
                        {getCollaboratedAgentName(agent)}
                      </span>
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 text-slate-600">
                      {agent.proprietor || "-"}
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 text-slate-600">
                      {agent.mobile || "-"}
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {agent.upazila}, {agent.district}
                    </td>
                    {canManage && (
                      <td className="border border-slate-300 px-3 py-1.5 text-center space-x-1.5 whitespace-nowrap">
                        <button
                          onClick={() => handleOpenEdit(agent)}
                          className="px-2 py-1 rounded bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-medium inline-flex items-center gap-1"
                        >
                          <FaEdit size={10} /> Edit
                        </button>
                        <button
                          onClick={() => handleDelete(agent)}
                          className="px-2 py-1 rounded bg-red-600 hover:bg-red-700 text-white text-[11px] font-medium inline-flex items-center gap-1"
                        >
                          <FaTrash size={10} /> Delete
                        </button>
                      </td>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add / Edit Modal */}
      {isModalOpen && canManage && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-3">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md max-h-[90vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between px-4 py-3 border-b border-slate-200 bg-slate-100">
              <h2 className="text-sm font-bold text-slate-800">
                {editingAgent ? "Edit DOC Agent" : "Add New DOC Agent"}
              </h2>
              <button
                onClick={() => setIsModalOpen(false)}
                className="w-7 h-7 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes />
              </button>
            </div>

            <form
              onSubmit={handleSave}
              className="p-4 space-y-3 overflow-y-auto text-xs"
            >
              {/* 1. District & Upazila First so Upazila Agent List Filters Automatically */}
              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    District *
                  </label>
                  <select
                    required
                    value={formData.district}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        district: e.target.value,
                        upazila: "",
                        feed_agent_id: "",
                      })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none bg-white"
                  >
                    <option value="">Select District</option>
                    {allowedDistricts.map((d) => (
                      <option key={d.id} value={d.name}>
                        {d.name}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Upazila *
                  </label>
                  <select
                    required
                    disabled={!formData.district}
                    value={formData.upazila}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        upazila: e.target.value,
                        feed_agent_id: "",
                      })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none bg-white disabled:bg-slate-100"
                  >
                    <option value="">Select Upazila</option>
                    {allowedUpazilasForSelectedDistrict.map((u) => (
                      <option key={u.id} value={u.name}>
                        {u.name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* 2. Collaborate with Agent from `agents` Table of Selected Upazila */}
              <div className="bg-indigo-50/70 border border-indigo-200 rounded-xl p-2.5">
                <label className="block font-bold text-indigo-900 mb-1">
                  🔗 Collaborate with Agent ({formData.upazila || "Select Upazila"})
                </label>
                <select
                  value={formData.feed_agent_id}
                  onChange={(e) =>
                    handleSelectCollaboratedAgent(e.target.value)
                  }
                  className="w-full border border-indigo-300 rounded-lg px-2.5 py-1.5 outline-none bg-white text-slate-800"
                >
                  <option value="">
                    -- Select Agent from Agents List (
                    {upazilaAgentsForCollaboration.length}) --
                  </option>
                  {upazilaAgentsForCollaboration.map((fa) => (
                    <option key={fa.id} value={String(fa.id)}>
                      {fa.agent_name}
                      {fa.agent_code ? ` [ID: ${fa.agent_code}]` : ""} (
                      {fa.companies?.company_name || "Agent"})
                    </option>
                  ))}
                </select>
              </div>

              {/* 3. DOC.ID & Agent Name */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    DOC.ID *
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="e.g. 1059"
                    value={formData.doc_id}
                    onChange={(e) =>
                      setFormData({ ...formData, doc_id: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                  />
                </div>

                <div className="sm:col-span-2">
                  <label className="block font-semibold text-slate-700 mb-1">
                    DOC Agent Name *
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="Enter agent name"
                    value={formData.agent_name}
                    onChange={(e) =>
                      setFormData({ ...formData, agent_name: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2.5">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Proprietor
                  </label>
                  <input
                    type="text"
                    placeholder="Proprietor name"
                    value={formData.proprietor}
                    onChange={(e) =>
                      setFormData({ ...formData, proprietor: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Mobile
                  </label>
                  <input
                    type="text"
                    placeholder="017XXXXXXXX"
                    value={formData.mobile}
                    onChange={(e) =>
                      setFormData({ ...formData, mobile: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="flex-1 py-2 rounded-lg border border-slate-300 text-slate-600 font-semibold"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  disabled={submitting}
                  className="flex-1 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-semibold flex items-center justify-center gap-1.5"
                >
                  {submitting && <FaSpinner className="animate-spin" />}
                  {editingAgent ? "Save Changes" : "Add Agent"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}