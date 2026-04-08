import { useState, useEffect, useRef, useMemo } from "react";
import * as d3 from "d3";

// ============================================================
// CORTEX UI PROTOTYPE - Knowledge Graph + Notion-style Reports
// ============================================================

const COLORS = {
  bg: "#191919",
  surface: "#1e1e1e",
  card: "#252525",
  border: "#333",
  text: "#e0e0e0",
  textMuted: "#888",
  textDim: "#555",
  accent: "#5b9bd5",
  accentGlow: "rgba(91,155,213,0.3)",
  green: "#4caf50",
  red: "#ef5350",
  orange: "#ff9800",
  purple: "#9c6ade",
  yellow: "#ffd54f",
  cyan: "#4dd0e1",
  pink: "#f06292",
};

const nodeColors = {
  project: "#5b9bd5",
  costCode: "#4caf50",
  worker: "#ff9800",
  vendor: "#9c6ade",
  document: "#4dd0e1",
  gc: "#f06292",
  phase: "#ffd54f",
  insight: "#ef5350",
};

// Knowledge graph data for OWP Job 2012
const graphData = {
  nodes: [
    { id: "owp2012", label: "Job #2012\nExxel 8th Ave", type: "project", size: 28 },
    // Cost codes
    { id: "cc120", label: "120\nRoughin", type: "costCode", size: 22 },
    { id: "cc130", label: "130\nFinish", type: "costCode", size: 16 },
    { id: "cc111", label: "111\nGarage", type: "costCode", size: 14 },
    { id: "cc141", label: "141\nWater Main", type: "costCode", size: 13 },
    { id: "cc142", label: "142\nMech Room", type: "costCode", size: 12 },
    { id: "cc145", label: "145\nTub/Shower", type: "costCode", size: 11 },
    { id: "cc100", label: "100\nSupervision", type: "costCode", size: 10 },
    { id: "cc601", label: "601\nEngineering", type: "costCode", size: 10 },
    // Workers
    { id: "w_gerard", label: "Gerard\nJeffrey S", type: "worker", size: 14 },
    { id: "w_quint", label: "Quintanilla\nEsteban R", type: "worker", size: 13 },
    { id: "w_veley", label: "Veley\nNathaniel S", type: "worker", size: 12 },
    { id: "w_palma", label: "Palma Vides\nHugo", type: "worker", size: 12 },
    { id: "w_paco", label: "Paco Leyva\nOrlando", type: "worker", size: 11 },
    { id: "w_cortes", label: "Cortes M.\nVictor H", type: "worker", size: 11 },
    { id: "w_meza", label: "Meza Fuentes\nErick A", type: "worker", size: 11 },
    { id: "w_sanders", label: "Sanders\nAllen O", type: "worker", size: 11 },
    // Vendors
    { id: "v_rosen", label: "Rosen\nSupply Co", type: "vendor", size: 14 },
    { id: "v_ferguson", label: "Ferguson\nEnterprises", type: "vendor", size: 13 },
    { id: "v_keller", label: "Keller\nSupply", type: "vendor", size: 10 },
    { id: "v_franklin", label: "Franklin\nEngineering", type: "vendor", size: 11 },
    { id: "v_manor", label: "Manor\nHardware", type: "vendor", size: 9 },
    // GC
    { id: "gc_exxel", label: "Exxel\nPacific", type: "gc", size: 16 },
    // Documents
    { id: "d_jcr", label: "Job Detail\nReport", type: "document", size: 15 },
    { id: "d_contract", label: "Contract\n$1.39M", type: "document", size: 12 },
    { id: "d_payapps", label: "Pay Apps\n(8 invoices)", type: "document", size: 11 },
    { id: "d_lien", label: "Lien\nReleases", type: "document", size: 10 },
    // Insights
    { id: "i_roughin", label: "⚠️ Roughin\n+27% Over", type: "insight", size: 14 },
    { id: "i_mech", label: "🔴 Mech Room\n+178% Over", type: "insight", size: 13 },
    { id: "i_watermain", label: "✅ Water Main\n62% Under", type: "insight", size: 12 },
  ],
  links: [
    // Project connections
    { source: "owp2012", target: "gc_exxel", strength: 0.8 },
    { source: "owp2012", target: "d_jcr", strength: 0.9 },
    { source: "owp2012", target: "d_contract", strength: 0.7 },
    { source: "owp2012", target: "d_payapps", strength: 0.6 },
    { source: "owp2012", target: "d_lien", strength: 0.5 },
    // Cost codes to project
    { source: "owp2012", target: "cc120", strength: 0.9 },
    { source: "owp2012", target: "cc130", strength: 0.7 },
    { source: "owp2012", target: "cc111", strength: 0.6 },
    { source: "owp2012", target: "cc141", strength: 0.6 },
    { source: "owp2012", target: "cc142", strength: 0.5 },
    { source: "owp2012", target: "cc145", strength: 0.5 },
    { source: "owp2012", target: "cc100", strength: 0.5 },
    { source: "owp2012", target: "cc601", strength: 0.4 },
    // Workers to cost codes
    { source: "w_gerard", target: "cc100", strength: 0.9 },
    { source: "w_gerard", target: "cc120", strength: 0.3 },
    { source: "w_quint", target: "cc120", strength: 0.9 },
    { source: "w_quint", target: "cc142", strength: 0.6 },
    { source: "w_veley", target: "cc120", strength: 0.8 },
    { source: "w_veley", target: "cc145", strength: 0.4 },
    { source: "w_palma", target: "cc120", strength: 0.8 },
    { source: "w_paco", target: "cc120", strength: 0.7 },
    { source: "w_paco", target: "cc111", strength: 0.4 },
    { source: "w_cortes", target: "cc120", strength: 0.7 },
    { source: "w_cortes", target: "cc111", strength: 0.4 },
    { source: "w_meza", target: "cc130", strength: 0.9 },
    { source: "w_sanders", target: "cc130", strength: 0.8 },
    // Vendors to cost codes
    { source: "v_rosen", target: "cc120", strength: 0.7 },
    { source: "v_rosen", target: "cc111", strength: 0.4 },
    { source: "v_ferguson", target: "cc120", strength: 0.6 },
    { source: "v_ferguson", target: "cc142", strength: 0.3 },
    { source: "v_keller", target: "cc120", strength: 0.3 },
    { source: "v_franklin", target: "cc601", strength: 0.9 },
    { source: "v_manor", target: "cc145", strength: 0.7 },
    // GC connections
    { source: "gc_exxel", target: "d_contract", strength: 0.8 },
    { source: "gc_exxel", target: "d_payapps", strength: 0.7 },
    // Insights
    { source: "i_roughin", target: "cc120", strength: 0.9 },
    { source: "i_mech", target: "cc142", strength: 0.9 },
    { source: "i_watermain", target: "cc141", strength: 0.9 },
    // JCR connections
    { source: "d_jcr", target: "cc120", strength: 0.5 },
    { source: "d_jcr", target: "cc130", strength: 0.4 },
    { source: "d_jcr", target: "cc111", strength: 0.3 },
  ],
};

// ============================================================
// GRAPH VIEW COMPONENT
// ============================================================
function GraphView({ onNodeClick, selectedNode }) {
  const svgRef = useRef(null);
  const simRef = useRef(null);

  useEffect(() => {
    if (!svgRef.current) return;
    const svg = d3.select(svgRef.current);
    svg.selectAll("*").remove();

    const width = svgRef.current.clientWidth;
    const height = svgRef.current.clientHeight;

    const g = svg.append("g");

    // Zoom
    const zoom = d3.zoom().scaleExtent([0.3, 4]).on("zoom", (e) => g.attr("transform", e.transform));
    svg.call(zoom);
    svg.call(zoom.transform, d3.zoomIdentity.translate(width / 2, height / 2).scale(0.8));

    // Glow filter
    const defs = svg.append("defs");
    const filter = defs.append("filter").attr("id", "glow");
    filter.append("feGaussianBlur").attr("stdDeviation", "4").attr("result", "coloredBlur");
    const merge = filter.append("feMerge");
    merge.append("feMergeNode").attr("in", "coloredBlur");
    merge.append("feMergeNode").attr("in", "SourceGraphic");

    const nodes = graphData.nodes.map((d) => ({ ...d }));
    const links = graphData.links.map((d) => ({ ...d }));

    const sim = d3.forceSimulation(nodes)
      .force("link", d3.forceLink(links).id((d) => d.id).distance(100).strength((d) => d.strength * 0.3))
      .force("charge", d3.forceManyBody().strength(-300))
      .force("center", d3.forceCenter(0, 0))
      .force("collision", d3.forceCollide().radius((d) => d.size + 10));

    simRef.current = sim;

    // Links
    const link = g.append("g").selectAll("line").data(links).join("line")
      .attr("stroke", "rgba(91,155,213,0.15)")
      .attr("stroke-width", (d) => d.strength * 2.5);

    // Node groups
    const node = g.append("g").selectAll("g").data(nodes).join("g")
      .style("cursor", "pointer")
      .call(d3.drag()
        .on("start", (e, d) => { if (!e.active) sim.alphaTarget(0.3).restart(); d.fx = d.x; d.fy = d.y; })
        .on("drag", (e, d) => { d.fx = e.x; d.fy = e.y; })
        .on("end", (e, d) => { if (!e.active) sim.alphaTarget(0); d.fx = null; d.fy = null; })
      );

    // Circles
    node.append("circle")
      .attr("r", (d) => d.size)
      .attr("fill", (d) => nodeColors[d.type])
      .attr("fill-opacity", 0.8)
      .attr("stroke", (d) => nodeColors[d.type])
      .attr("stroke-width", 2)
      .attr("stroke-opacity", 0.4)
      .style("filter", "url(#glow)");

    // Labels
    node.each(function (d) {
      const lines = d.label.split("\n");
      const el = d3.select(this);
      lines.forEach((line, i) => {
        el.append("text")
          .attr("text-anchor", "middle")
          .attr("dy", i === 0 ? (lines.length > 1 ? -5 : 4) : 10)
          .attr("fill", "#fff")
          .attr("font-size", d.size > 14 ? "9px" : "7px")
          .attr("font-family", "Inter, system-ui, sans-serif")
          .attr("font-weight", i === 0 ? "600" : "400")
          .attr("pointer-events", "none")
          .text(line);
      });
    });

    // Hover & click
    node.on("mouseover", function (e, d) {
      d3.select(this).select("circle").transition().duration(200).attr("r", d.size * 1.3).attr("fill-opacity", 1);
      link.transition().duration(200)
        .attr("stroke", (l) => (l.source.id === d.id || l.target.id === d.id) ? nodeColors[d.type] : "rgba(91,155,213,0.06)")
        .attr("stroke-width", (l) => (l.source.id === d.id || l.target.id === d.id) ? 3 : 0.5)
        .attr("stroke-opacity", (l) => (l.source.id === d.id || l.target.id === d.id) ? 0.8 : 0.3);
    }).on("mouseout", function (e, d) {
      d3.select(this).select("circle").transition().duration(200).attr("r", d.size).attr("fill-opacity", 0.8);
      link.transition().duration(200).attr("stroke", "rgba(91,155,213,0.15)").attr("stroke-width", (l) => l.strength * 2.5).attr("stroke-opacity", 1);
    }).on("click", (e, d) => onNodeClick?.(d));

    sim.on("tick", () => {
      link.attr("x1", (d) => d.source.x).attr("y1", (d) => d.source.y).attr("x2", (d) => d.target.x).attr("y2", (d) => d.target.y);
      node.attr("transform", (d) => `translate(${d.x},${d.y})`);
    });

    return () => sim.stop();
  }, []);

  return (
    <div style={{ width: "100%", height: "100%", position: "relative" }}>
      <svg ref={svgRef} style={{ width: "100%", height: "100%", background: COLORS.bg }} />
      {/* Legend */}
      <div style={{ position: "absolute", bottom: 16, left: 16, background: "rgba(25,25,25,0.9)", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "12px 16px", display: "flex", gap: 16, flexWrap: "wrap" }}>
        {Object.entries(nodeColors).map(([type, color]) => (
          <div key={type} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div style={{ width: 10, height: 10, borderRadius: "50%", background: color }} />
            <span style={{ color: COLORS.textMuted, fontSize: 11, textTransform: "capitalize" }}>{type === "costCode" ? "Cost Code" : type === "gc" ? "GC" : type}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ============================================================
// NOTION-STYLE REPORT COMPONENTS
// ============================================================
function NotionBlock({ children, style }) {
  return <div style={{ marginBottom: 2, padding: "3px 4px", borderRadius: 4, ...style }}>{children}</div>;
}

function NotionH1({ children }) {
  return <h1 style={{ fontSize: 28, fontWeight: 700, color: COLORS.text, margin: "24px 0 8px", fontFamily: "Inter, system-ui, sans-serif", letterSpacing: "-0.02em" }}>{children}</h1>;
}

function NotionH2({ children }) {
  return <h2 style={{ fontSize: 20, fontWeight: 600, color: COLORS.text, margin: "20px 0 6px", fontFamily: "Inter, system-ui, sans-serif" }}>{children}</h2>;
}

function NotionH3({ children }) {
  return <h3 style={{ fontSize: 16, fontWeight: 600, color: COLORS.text, margin: "16px 0 4px", fontFamily: "Inter, system-ui, sans-serif" }}>{children}</h3>;
}

function NotionText({ children, muted }) {
  return <p style={{ fontSize: 14, lineHeight: 1.7, color: muted ? COLORS.textMuted : COLORS.text, margin: "4px 0", fontFamily: "Inter, system-ui, sans-serif" }}>{children}</p>;
}

function NotionCallout({ emoji, children, color = "#2a2a2a", borderColor = COLORS.border }) {
  return (
    <div style={{ display: "flex", gap: 12, padding: "16px 16px", background: color, border: `1px solid ${borderColor}`, borderRadius: 6, margin: "8px 0" }}>
      <span style={{ fontSize: 20, flexShrink: 0 }}>{emoji}</span>
      <div style={{ flex: 1 }}>{children}</div>
    </div>
  );
}

function NotionTable({ headers, rows }) {
  return (
    <div style={{ overflowX: "auto", margin: "8px 0", borderRadius: 6, border: `1px solid ${COLORS.border}` }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "Inter, system-ui, sans-serif" }}>
        <thead>
          <tr style={{ background: "#2a2a2a" }}>
            {headers.map((h, i) => (
              <th key={i} style={{ padding: "10px 14px", textAlign: i === 0 ? "left" : "right", color: COLORS.textMuted, fontWeight: 500, borderBottom: `1px solid ${COLORS.border}`, whiteSpace: "nowrap", fontSize: 12, textTransform: "uppercase", letterSpacing: "0.05em" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, ri) => (
            <tr key={ri} style={{ background: ri % 2 === 0 ? "transparent" : "rgba(255,255,255,0.02)" }}>
              {row.map((cell, ci) => (
                <td key={ci} style={{ padding: "10px 14px", textAlign: ci === 0 ? "left" : "right", color: typeof cell === "object" ? cell.color : COLORS.text, borderBottom: `1px solid rgba(255,255,255,0.04)`, fontWeight: typeof cell === "object" && cell.bold ? 600 : 400, fontFamily: ci > 0 ? "SF Mono, Menlo, monospace" : "inherit", fontSize: ci > 0 ? 12 : 13 }}>
                  {typeof cell === "object" ? cell.value : cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function MetricCard({ label, value, subtext, trend, color = COLORS.accent }) {
  return (
    <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "20px 20px", flex: "1 1 180px", minWidth: 160 }}>
      <div style={{ fontSize: 11, color: COLORS.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8, fontFamily: "Inter, system-ui, sans-serif" }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 700, color, fontFamily: "Inter, system-ui, sans-serif", letterSpacing: "-0.02em" }}>{value}</div>
      {subtext && <div style={{ fontSize: 12, color: COLORS.textMuted, marginTop: 4 }}>{subtext}</div>}
      {trend && <div style={{ fontSize: 12, color: trend > 0 ? COLORS.green : COLORS.red, marginTop: 4 }}>{trend > 0 ? "↑" : "↓"} {Math.abs(trend)}% vs budget</div>}
    </div>
  );
}

function ProgressBar({ label, value, max, color = COLORS.accent }) {
  const pct = (value / max) * 100;
  const overBudget = pct > 100;
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
        <span style={{ fontSize: 13, color: COLORS.text, fontFamily: "Inter, system-ui, sans-serif" }}>{label}</span>
        <span style={{ fontSize: 12, color: overBudget ? COLORS.red : COLORS.textMuted, fontFamily: "SF Mono, Menlo, monospace" }}>{Math.round(pct)}%</span>
      </div>
      <div style={{ height: 6, background: "rgba(255,255,255,0.06)", borderRadius: 3, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${Math.min(pct, 100)}%`, background: overBudget ? COLORS.red : color, borderRadius: 3, transition: "width 0.5s ease" }} />
      </div>
    </div>
  );
}

// ============================================================
// REPORT VIEW
// ============================================================
function ReportView() {
  return (
    <div style={{ maxWidth: 860, margin: "0 auto", padding: "32px 48px", fontFamily: "Inter, system-ui, sans-serif" }}>
      {/* Cover */}
      <div style={{ marginBottom: 32 }}>
        <div style={{ fontSize: 11, color: COLORS.accent, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8 }}>Cortex Intelligence Report</div>
        <NotionH1>OWP Job #2012 — Exxel 8th Ave Apartments</NotionH1>
        <NotionText muted>163-unit mid-rise multifamily · Division 22 Plumbing · Seattle, WA · Jan 2013 – Jun 2014</NotionText>
        <div style={{ height: 1, background: COLORS.border, margin: "16px 0" }} />
      </div>

      {/* KPI Cards */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <MetricCard label="Contract Value" value="$1.39M" subtext="163 units" />
        <MetricCard label="JCR Net Profit" value="$533K" subtext="38.3% margin" color={COLORS.green} />
        <MetricCard label="Total Labor Hours" value="14,607" subtext="~90 hrs/unit" color={COLORS.orange} />
        <MetricCard label="Labor:Material" value="1.05:1" subtext="$440K : $418K" color={COLORS.purple} />
      </div>

      {/* Budget vs Actual */}
      <NotionH2>Budget vs Actual by Phase</NotionH2>
      <NotionText muted>Red indicates over budget. Green indicates savings.</NotionText>

      <NotionTable
        headers={["Phase", "Budget", "Actual", "Variance", "Hours", "$/Hr"]}
        rows={[
          ["120 · Roughin Labor", "$139,800", "$177,292", { value: "+$37,492", color: COLORS.red, bold: true }, "9,320", "$19.02"],
          ["130 · Finish Labor", "$31,000", "$34,658", { value: "+$3,658", color: COLORS.red }, "1,841", "$18.83"],
          ["111 · Garage Labor", "$14,600", "$19,354", { value: "+$4,754", color: COLORS.red }, "872", "$22.19"],
          ["142 · Mech Room", "$3,525", "$9,784", { value: "+$6,259", color: COLORS.red, bold: true }, "379", "$25.82"],
          ["141 · Water Main", "$41,000", "$15,756", { value: "-$25,244", color: COLORS.green, bold: true }, "800", "$19.70"],
          ["145 · Tub/Shower", "$19,800", "$6,781", { value: "-$13,019", color: COLORS.green }, "338", "$20.06"],
          ["112 · Canout", "$11,700", "$5,424", { value: "-$6,276", color: COLORS.green }, "238", "$22.79"],
          ["100 · Supervision", "$8,100", "$6,023", { value: "-$2,077", color: COLORS.green }, "165", "$36.50"],
        ]}
      />

      {/* Budget consumption bars */}
      <NotionH3>Budget Consumption</NotionH3>
      <div style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 20, margin: "8px 0" }}>
        <ProgressBar label="Roughin Labor" value={177292} max={139800} color={COLORS.red} />
        <ProgressBar label="Mech Room" value={9784} max={3525} color={COLORS.red} />
        <ProgressBar label="Garage Labor" value={19354} max={14600} color={COLORS.orange} />
        <ProgressBar label="Finish Labor" value={34658} max={31000} color={COLORS.orange} />
        <ProgressBar label="Supervision" value={6023} max={8100} color={COLORS.green} />
        <ProgressBar label="Water Main/Insulation" value={15756} max={41000} color={COLORS.accent} />
        <ProgressBar label="Tub/Shower" value={6781} max={19800} color={COLORS.accent} />
      </div>

      {/* Insights */}
      <NotionH2>Agent Insights</NotionH2>

      <NotionCallout emoji="🔴" color="rgba(239,83,80,0.08)" borderColor="rgba(239,83,80,0.2)">
        <div style={{ fontWeight: 600, fontSize: 14, color: COLORS.red, marginBottom: 4 }}>Roughin Labor — 27% Over Budget</div>
        <NotionText>9,320 hours consumed (64% of all labor). At 57 hrs/unit, this exceeds comparable projects by ~15%. Recommend budgeting 65+ hrs/unit on future mid-rise and adding 20% contingency.</NotionText>
      </NotionCallout>

      <NotionCallout emoji="🔴" color="rgba(239,83,80,0.08)" borderColor="rgba(239,83,80,0.2)">
        <div style={{ fontWeight: 600, fontSize: 14, color: COLORS.red, marginBottom: 4 }}>Mech Room — 178% Over Budget</div>
        <NotionText>$9,784 actual vs $3,525 budget. Complex manifold assemblies and circ pump installations consistently underestimated. Budget 2.5–3x initial estimate on future projects.</NotionText>
      </NotionCallout>

      <NotionCallout emoji="🟢" color="rgba(76,175,80,0.08)" borderColor="rgba(76,175,80,0.2)">
        <div style={{ fontWeight: 600, fontSize: 14, color: COLORS.green, marginBottom: 4 }}>Water Main/Insulation — 62% Under Budget</div>
        <NotionText>$15,756 actual vs $41,000 budget. Estimate is too conservative by ~$25K. Tighten by 40–50% on future bids to sharpen competitiveness without risk.</NotionText>
      </NotionCallout>

      <NotionCallout emoji="💡" color="rgba(91,155,213,0.08)" borderColor="rgba(91,155,213,0.2)">
        <div style={{ fontWeight: 600, fontSize: 14, color: COLORS.accent, marginBottom: 4 }}>Per-Unit Benchmark: $5,265 direct cost / $8,537 revenue</div>
        <NotionText>Use $5,000–$5,500/unit as base cost for similar Division 22 scope on mid-rise multifamily. Target $8,000–$9,000/unit revenue for healthy margin.</NotionText>
      </NotionCallout>

      {/* Per Unit */}
      <NotionH2>Per-Unit Benchmarks (163 Units)</NotionH2>
      <NotionTable
        headers={["Metric", "Total", "Per Unit", "% of Revenue"]}
        rows={[
          [{ value: "Revenue", bold: true }, "$1,391,455", "$8,537", "100%"],
          [{ value: "Total Direct Cost", bold: true }, "$858,181", "$5,265", "61.7%"],
          [{ value: "JCR Net Profit", bold: true }, "$533,274", { value: "$3,272", color: COLORS.green, bold: true }, "38.3%"],
          ["", "", "", ""],
          ["Roughin Labor", "$177,292", "$1,088", "12.7%"],
          ["Roughin Material", "$22,695", "$139", "1.6%"],
          ["Finish Labor", "$34,658", "$213", "2.5%"],
          ["Engineering", "$15,682", "$96", "1.1%"],
          ["Permits", "$9,742", "$60", "0.7%"],
        ]}
      />

      {/* Crew */}
      <NotionH2>Crew Rate Tiers</NotionH2>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
        {[
          { tier: "Superintendent", range: "$33–38/hr", count: 1, color: COLORS.pink },
          { tier: "Lead Journeyman", range: "$28–31/hr", count: 4, color: COLORS.orange },
          { tier: "Journeyman", range: "$20–27/hr", count: 8, color: COLORS.accent },
          { tier: "Apprentice", range: "$12–16/hr", count: 7, color: COLORS.cyan },
          { tier: "Helper", range: "$12–14/hr", count: 8, color: COLORS.purple },
        ].map((t) => (
          <div key={t.tier} style={{ background: COLORS.card, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "14px 18px", flex: "1 1 150px", minWidth: 140 }}>
            <div style={{ fontSize: 11, color: t.color, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>{t.tier}</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: COLORS.text, marginTop: 4 }}>{t.range}</div>
            <div style={{ fontSize: 12, color: COLORS.textMuted, marginTop: 2 }}>{t.count} workers</div>
          </div>
        ))}
      </div>

      <div style={{ height: 1, background: COLORS.border, margin: "32px 0" }} />
      <NotionText muted>Generated by Cortex Intelligence Agent · Analysis of 135-page Job Detail Report · {new Date().toLocaleDateString()}</NotionText>
    </div>
  );
}

// ============================================================
// NODE DETAIL PANEL
// ============================================================
function NodePanel({ node, onClose }) {
  if (!node) return null;
  const details = {
    owp2012: { title: "Job #2012 — Exxel 8th Ave", stats: [["Revenue", "$1,391,455"], ["Expenses", "$858,181"], ["Net", "$533,274"], ["Units", "163"], ["Duration", "17 months"]] },
    cc120: { title: "120 · Roughin Labor", stats: [["Budget", "$139,800"], ["Actual", "$177,292"], ["Variance", "+27%"], ["Hours", "9,320"], ["Workers", "15+"]] },
    cc142: { title: "142 · Mech Room", stats: [["Budget", "$3,525"], ["Actual", "$9,784"], ["Variance", "+178%"], ["Hours", "379"], ["Risk", "HIGH"]] },
    w_gerard: { title: "Gerard, Jeffrey S", stats: [["Role", "Superintendent"], ["Rate", "~$36.50/hr"], ["Hours (this job)", "165"], ["Phases", "Supervision, Roughin"]] },
    w_quint: { title: "Quintanilla, Esteban R", stats: [["Role", "Lead Journeyman"], ["Rate", "~$31/hr"], ["Phases", "Roughin, Mech Room, Garage"]] },
    v_rosen: { title: "Rosen Supply Company", stats: [["Type", "Primary Supplier"], ["Materials", "Pipe, fittings, valves"], ["Cost Codes", "210, 211, 220"]] },
    gc_exxel: { title: "Exxel Pacific, Inc.", stats: [["Role", "General Contractor"], ["Contract", "$1,394,655"], ["Retainage", "$69,573 (5%)"], ["Payment", "Monthly pay apps"]] },
    d_jcr: { title: "Job Detail Report", stats: [["Pages", "135"], ["Cost Codes", "20+"], ["Date Range", "Jan 2013–Jun 2014"], ["Source Types", "PR, AP, GL, AR"]] },
  };
  const info = details[node.id] || { title: node.label.replace("\n", " "), stats: [["Type", node.type]] };

  return (
    <div style={{ position: "absolute", right: 16, top: 16, width: 280, background: "rgba(30,30,30,0.95)", border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: 20, backdropFilter: "blur(20px)", zIndex: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div style={{ width: 10, height: 10, borderRadius: "50%", background: nodeColors[node.type] }} />
        <button onClick={onClose} style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer", fontSize: 18 }}>×</button>
      </div>
      <div style={{ fontSize: 16, fontWeight: 700, color: COLORS.text, marginBottom: 16, fontFamily: "Inter, system-ui, sans-serif" }}>{info.title}</div>
      {info.stats.map(([k, v], i) => (
        <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderBottom: `1px solid rgba(255,255,255,0.04)` }}>
          <span style={{ fontSize: 12, color: COLORS.textMuted }}>{k}</span>
          <span style={{ fontSize: 12, color: COLORS.text, fontWeight: 500 }}>{v}</span>
        </div>
      ))}
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
export default function CortexPrototype() {
  const [view, setView] = useState("graph");
  const [selectedNode, setSelectedNode] = useState(null);

  return (
    <div style={{ width: "100%", height: "100vh", background: COLORS.bg, color: COLORS.text, display: "flex", flexDirection: "column", fontFamily: "Inter, system-ui, sans-serif" }}>
      {/* Top Nav */}
      <div style={{ height: 52, borderBottom: `1px solid ${COLORS.border}`, display: "flex", alignItems: "center", padding: "0 20px", gap: 16, flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ width: 24, height: 24, borderRadius: 6, background: `linear-gradient(135deg, ${COLORS.accent}, ${COLORS.purple})`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <span style={{ fontSize: 13, fontWeight: 800, color: "#fff" }}>C</span>
          </div>
          <span style={{ fontSize: 15, fontWeight: 700, letterSpacing: "-0.01em" }}>Cortex</span>
        </div>
        <div style={{ width: 1, height: 24, background: COLORS.border }} />
        <div style={{ display: "flex", gap: 2 }}>
          {[
            { id: "graph", label: "Knowledge Graph", icon: "◉" },
            { id: "report", label: "Intelligence Report", icon: "◧" },
          ].map((tab) => (
            <button key={tab.id} onClick={() => setView(tab.id)}
              style={{
                padding: "6px 14px", border: "none", borderRadius: 6, cursor: "pointer",
                background: view === tab.id ? "rgba(91,155,213,0.15)" : "transparent",
                color: view === tab.id ? COLORS.accent : COLORS.textMuted,
                fontSize: 13, fontWeight: 500, fontFamily: "Inter, system-ui, sans-serif",
                transition: "all 0.15s ease",
              }}>
              {tab.icon} {tab.label}
            </button>
          ))}
        </div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: 12, color: COLORS.textMuted }}>OWP, LLC · Job #2012</div>
      </div>

      {/* Content */}
      <div style={{ flex: 1, position: "relative", overflow: "hidden" }}>
        {view === "graph" ? (
          <>
            <GraphView onNodeClick={setSelectedNode} selectedNode={selectedNode} />
            <NodePanel node={selectedNode} onClose={() => setSelectedNode(null)} />
            {/* Floating instructions */}
            <div style={{ position: "absolute", top: 16, left: 16, background: "rgba(30,30,30,0.85)", border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: "10px 14px", fontSize: 11, color: COLORS.textMuted, backdropFilter: "blur(10px)" }}>
              Drag nodes · Scroll to zoom · Click for details
            </div>
          </>
        ) : (
          <div style={{ height: "100%", overflow: "auto" }}>
            <ReportView />
          </div>
        )}
      </div>
    </div>
  );
}
