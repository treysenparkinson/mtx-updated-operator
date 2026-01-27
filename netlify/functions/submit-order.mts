import type { Context, Config } from "@netlify/functions";
import * as XLSX from 'xlsx';

export default async (req: Request, context: Context) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { status: 204 });
  }

  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  try {
    const data = await req.json();
    const { refId, contactName, contactEmail, labels } = data;

    if (!refId || !labels || labels.length === 0) {
      return new Response(JSON.stringify({ error: "Missing required fields" }), { status: 400 });
    }

    const timestamp = new Date().toISOString();
    const formattedDate = new Date().toLocaleString("en-US", {
      year: "numeric", month: "short", day: "2-digit",
      hour: "2-digit", minute: "2-digit"
    });

    // Build Excel data - one row per quantity
    const excelData: any[][] = [];
    
    // Row 1: Reference ID
    excelData.push([`Reference ID: ${refId}`]);
    
    // Row 2: Empty row
    excelData.push([]);
    
    // Row 3: Headers
    const headers = ["Size", "Color", "VAR1", "VAR2", "VAR3", "VAR4", "VAR5", "VAR6", "VAR1 Size", "VAR2 Size", "VAR3 Size", "VAR4 Size", "VAR5 Size", "VAR6 Size", "Font"];
    excelData.push(headers);
    
    let totalLabels = 0;
    labels.forEach((label: any) => {
      const row = [
        label.size?.name || "",
        label.color?.name || "",
        label.var1 || "",
        label.var2 || "",
        label.var3 || "",
        label.var4 || "",
        label.var5 || "",
        label.var6 || "",
        label.var1 ? (label.var1Size || 18) : "",
        label.var2 ? (label.var2Size || 18) : "",
        label.var3 ? (label.var3Size || 18) : "",
        label.var4 ? (label.var4Size || 10) : "",
        label.var5 ? (label.var5Size || 10) : "",
        label.var6 ? (label.var6Size || 10) : "",
        label.font?.name || "Calibri (Default)"
      ];
      
      const qty = label.quantity || 1;
      totalLabels += qty;
      for (let i = 0; i < qty; i++) {
        excelData.push(row);
      }
    });

    // Create XLSX workbook
    const worksheet = XLSX.utils.aoa_to_sheet(excelData);
    
    // Set column widths
    worksheet['!cols'] = [
      { wch: 15 }, // Size
      { wch: 12 }, // Color
      { wch: 15 }, // VAR1
      { wch: 15 }, // VAR2
      { wch: 15 }, // VAR3
      { wch: 15 }, // VAR4
      { wch: 15 }, // VAR5
      { wch: 15 }, // VAR6
      { wch: 10 }, // VAR1 Size
      { wch: 10 }, // VAR2 Size
      { wch: 10 }, // VAR3 Size
      { wch: 10 }, // VAR4 Size
      { wch: 10 }, // VAR5 Size
      { wch: 10 }, // VAR6 Size
      { wch: 18 }, // Font
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Labels");
    
    // Generate XLSX as base64
    const xlsxBuffer = XLSX.write(workbook, { type: 'base64', bookType: 'xlsx' });

    // Generate PDF HTML
    const generateLabelSVG = (label: any) => {
      const size = label.size || { width: 160, height: 182, id: "30mm-standard", name: "30MM Standard" };
      const color = label.color || { bg: "#16a34a", text: "#fff", id: "green-white" };
      const corners = label.corners || "squared";
      const font = label.font || { family: "Calibri, sans-serif" };
      const positions = label.positions || {};
      
      const scale = 1.5;
      const w = size.width * scale;
      const h = size.height * scale;
      const cutoutR = (size.id === "22mm" ? 24 : 36) * scale;
      const cutoutY = h * 0.68;
      const borderR = corners === "rounded" ? 8 * scale : 0;
      const isWhiteBlack = color.id === "white-black";
      
      const showLine2 = size.id !== "30mm-short";
      const showLine3 = size.id !== "22mm" && size.id !== "30mm-short";

      const defPos = {
        var1: { x: size.width / 2, y: 20 },
        var2: { x: size.width / 2, y: 38 },
        var3: { x: size.width / 2, y: 54 },
        var4: { x: size.width / 2, y: size.height * 0.68 - (size.id === "22mm" ? 24 : 36) - 16 },
        var5: { x: size.width / 2 - (size.id === "22mm" ? 28 : 50), y: size.height * 0.68 - 28 },
        var6: { x: size.width / 2 + (size.id === "22mm" ? 28 : 50), y: size.height * 0.68 - 28 }
      };
      const pos = { ...defPos, ...positions };

      let textElements = "";
      if (label.var1) textElements += `<text x="${pos.var1.x * scale}" y="${pos.var1.y * scale}" text-anchor="middle" fill="${color.text}" font-size="${(label.var1Size || 18) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var1)}</text>`;
      if (label.var2 && showLine2) textElements += `<text x="${pos.var2.x * scale}" y="${pos.var2.y * scale}" text-anchor="middle" fill="${color.text}" font-size="${(label.var2Size || 18) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var2)}</text>`;
      if (label.var3 && showLine3) textElements += `<text x="${pos.var3.x * scale}" y="${pos.var3.y * scale}" text-anchor="middle" fill="${color.text}" font-size="${(label.var3Size || 18) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var3)}</text>`;
      if (label.var4) textElements += `<text x="${pos.var4.x * scale}" y="${pos.var4.y * scale}" text-anchor="middle" fill="${color.text}" font-size="${(label.var4Size || 10) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var4)}</text>`;
      if (label.var5) textElements += `<text x="${pos.var5.x * scale}" y="${pos.var5.y * scale}" text-anchor="end" fill="${color.text}" font-size="${(label.var5Size || 10) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var5)}</text>`;
      if (label.var6) textElements += `<text x="${pos.var6.x * scale}" y="${pos.var6.y * scale}" text-anchor="start" fill="${color.text}" font-size="${(label.var6Size || 10) * scale * 0.7}" font-family="${font.family || 'Calibri, sans-serif'}">${escapeHtml(label.var6)}</text>`;

      const borderStroke = isWhiteBlack ? `stroke="#000" stroke-width="2"` : "";
      const circleStroke = isWhiteBlack ? `stroke="#000" stroke-width="2"` : `stroke="${color.bg}" stroke-width="4"`;
      
      return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">
        <rect width="${w}" height="${h}" fill="${color.bg}" rx="${borderR}" ${borderStroke}/>
        <circle cx="${w/2}" cy="${cutoutY}" r="${cutoutR}" fill="#ffffff" ${circleStroke}/>
        ${textElements}
      </svg>`;
    };

    const escapeHtml = (text: string) => {
      return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    };

    const labelCards = labels.map((label: any) => `
      <div style="display: flex; align-items: center; padding: 20px; border-bottom: 1px solid #e2e8f0; gap: 20px;">
        <div style="flex-shrink: 0;">
          ${generateLabelSVG(label)}
        </div>
        <div style="flex: 1;">
          <div style="font-size: 18px; font-weight: 600; color: #1e293b;">${label.size?.name || "30MM Standard"}</div>
          <div style="font-size: 14px; color: #64748b;">${label.size?.dimensions || '2" × 2.27"'}</div>
          <div style="font-size: 14px; color: #64748b; margin-top: 4px;">Color: ${label.color?.name || "Green/White"}</div>
          <div style="font-size: 14px; color: #64748b;">Font: ${label.font?.name || "Calibri (Default)"}</div>
          <div style="font-size: 14px; color: #64748b;">Corners: ${label.corners || "squared"}</div>
          <div style="font-size: 14px; color: #64748b;">Notch: ${label.notch || "none"}</div>
        </div>
        <div style="font-size: 24px; font-weight: 700; color: #1e293b;">×${label.quantity || 1}</div>
      </div>
    `).join("");

    const pdfHtml = `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body { font-family: Arial, sans-serif; margin: 0; padding: 40px; background: #fff; }
.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; padding-bottom: 20px; border-bottom: 2px solid #1e293b; }
.ref-id { font-size: 14px; font-weight: 600; color: #1e293b; }
.date { font-size: 14px; color: #64748b; }
.title { font-size: 28px; font-weight: 700; color: #1e293b; margin-bottom: 10px; }
.contact { font-size: 14px; color: #64748b; margin-bottom: 20px; }
.labels-container { border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; }
.summary { margin-top: 20px; padding: 15px; background: #f8fafc; border-radius: 8px; }
.summary-text { font-size: 14px; color: #64748b; }
.summary-total { font-size: 18px; font-weight: 600; color: #1e293b; }
</style>
</head>
<body>
<div class="header">
<div class="ref-id">Reference ID: ${refId}</div>
<div class="date">${formattedDate}</div>
</div>
<div class="title">Saved Labels Summary</div>
<div class="contact">Contact: ${contactName || "N/A"} | Email: ${contactEmail || "N/A"}</div>
<div class="labels-container">
${labelCards}
</div>
<div class="summary">
<span class="summary-text">Total Labels: </span>
<span class="summary-total">${totalLabels}</span>
</div>
</body>
</html>`;

    const labelSummaries = labels.map((label: any) => ({
      size: label.size?.name,
      dimensions: label.size?.dimensions,
      color: label.color?.name,
      var1: label.var1 || "",
      var2: label.var2 || "",
      var3: label.var3 || "",
      var4: label.var4 || "",
      var5: label.var5 || "",
      var6: label.var6 || "",
      font: label.font?.name,
      corners: label.corners,
      notch: label.notch,
      quantity: label.quantity
    }));

    const webhookUrl = Netlify.env.get("ZAPIER_WEBHOOK_URL") || "https://hooks.zapier.com/hooks/catch/24455310/uqnnrvn/";
    
    const webhookPayload = {
      refId,
      contactName: contactName || "",
      contactEmail: contactEmail || "",
      timestamp,
      formattedDate,
      totalLabels,
      labelCount: labels.length,
      xlsxBase64: xlsxBuffer,
      xlsxFileName: `labels-${refId}.xlsx`,
      pdfHtml,
      pdfFileName: `labels-${refId}.pdf`,
      labels: labelSummaries
    };

    const webhookResponse = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(webhookPayload)
    });

    if (!webhookResponse.ok) {
      console.error("Webhook failed:", await webhookResponse.text());
    }

    return new Response(JSON.stringify({ 
      success: true, 
      message: "Order submitted successfully",
      refId,
      totalLabels
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" }
    });

  } catch (error) {
    console.error("Error processing order:", error);
    return new Response(JSON.stringify({ error: "Failed to process order" }), { status: 500 });
  }
};

export const config: Config = {
  path: "/api/submit-order"
};