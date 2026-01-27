import type { Context, Config } from "@netlify/functions";

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

    // Build Excel rows - one row per quantity
    const excelRows: string[][] = [];
    const headers = ["Size", "Color", "VAR1", "VAR2", "VAR3", "VAR4", "VAR5", "VAR6", "VAR1 Size", "VAR2 Size", "VAR3 Size", "VAR4 Size", "VAR5 Size", "VAR6 Size", "Font"];
    
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
        label.var1 ? String(label.var1Size || 18) : "",
        label.var2 ? String(label.var2Size || 18) : "",
        label.var3 ? String(label.var3Size || 18) : "",
        label.var4 ? String(label.var4Size || 10) : "",
        label.var5 ? String(label.var5Size || 10) : "",
        label.var6 ? String(label.var6Size || 10) : "",
        label.font?.name || "Calibri (Default)"
      ];
      
      const qty = label.quantity || 1;
      totalLabels += qty;
      for (let i = 0; i < qty; i++) {
        excelRows.push(row);
      }
    });

    // Generate CSV content
    const csvContent = [
      [`Reference ID: ${refId}`],
      [],
      headers,
      ...excelRows
    ].map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(",")).join("\n");

    // Build label summaries for webhook
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

    // Send to Zapier webhook
    const webhookUrl = Netlify.env.get("ZAPIER_WEBHOOK_URL") || "https://hooks.zapier.com/hooks/catch/24455310/uqnwsha/";
    
    const webhookPayload = {
      refId,
      contactName: contactName || "",
      contactEmail: contactEmail || "",
      timestamp,
      formattedDate,
      totalLabels,
      labelCount: labels.length,
      csvContent,
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