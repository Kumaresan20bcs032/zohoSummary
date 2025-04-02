require("dotenv").config();
const express = require("express");
const axios = require("axios");
const cors = require("cors");
const qs = require("qs");
const { formatDate } = require("./formatDate");

const app = express();
const PORT = process.env.PORT || 3000;

const ZOHO_API_URL = "https://www.zohoapis.in/crm/v2/Events";
const OUTLOOK_API_URL = "https://graph.microsoft.com/v1.0/me/calendar/events";

// Middleware
app.use(cors());
app.use(express.json());

// Token Cache
const tokenCache = {
  microsoft: { accessToken: null, expiresAt: null },
  zoho: { accessToken: null, expiresAt: null }
};

// Utility: Retry with Backoff (Handles API Rate Limits)
async function retryWithBackoff(fn, maxRetries = 5, delay = 5000) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error) {
      const status = error.response?.status || null;

      // If rate limited, use retry-after or exponential backoff
      if (status === 429) {
        const retryAfter = error.response?.headers['retry-after'];
        const waitTime = retryAfter ? parseInt(retryAfter) * 5000 : delay * Math.pow(2, i);
        console.warn(`⚠️ Rate limit hit. Retrying in ${waitTime / 5000} seconds...`);
        await new Promise(res => setTimeout(res, waitTime));
      } else {
        console.error(`❌ Error: ${error.message}`);
        if (i === maxRetries - 1) throw error;
      }
    }
  }
}

// Function to Get Microsoft Access Token
async function getMicrosoftAccessToken() {
  const now = Date.now();
  if (tokenCache.microsoft.accessToken && tokenCache.microsoft.expiresAt > now) {
    return tokenCache.microsoft.accessToken;
  }

  return retryWithBackoff(async () => {
    const response = await axios.post(
      `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      qs.stringify({
        client_id: process.env.MICROSOFT_CLIENT_ID,
        client_secret: process.env.MICROSOFT_CLIENT_SECRET,
        refresh_token: process.env.MICROSOFT_REFRESH_TOKEN,
        grant_type: "refresh_token",
        scope: "https://graph.microsoft.com/.default"
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    tokenCache.microsoft = {
      accessToken: response.data.access_token,
      expiresAt: now + (response.data.expires_in - 300) * 5000
    };

    return response.data.access_token;
  });
}

// Function to Get Zoho Access Token
async function getZohoAccessToken() {
  if (tokenCache.zoho.accessToken && tokenCache.zoho.expiresAt > Date.now()) {
    return tokenCache.zoho.accessToken;
  }

  return retryWithBackoff(async () => {
    const response = await axios.post("https://accounts.zoho.in/oauth/v2/token", qs.stringify({
      refresh_token: process.env.ZOHO_REFRESH_TOKEN,
      client_id: process.env.ZOHO_CLIENT_ID,
      client_secret: process.env.ZOHO_CLIENT_SECRET,
      grant_type: "refresh_token"
    }), { headers: { "Content-Type": "application/x-www-form-urlencoded" } });

    tokenCache.zoho = {
      accessToken: response.data.access_token,
      expiresAt: Date.now() + 3600 * 5000
    };

    return response.data.access_token;
  });
}

// Function to Fetch Zoho Events
async function fetchZohoEvents() {
  const token = await getZohoAccessToken();
  return retryWithBackoff(async () => {
    const response = await axios.get(ZOHO_API_URL, {
      headers: { Authorization: `Zoho-oauthtoken ${token}` }
    });
    return response.data.data;
  });
}

// Function to Create Outlook Events in Bulk
async function createOutlookEvents(events) {
  if (!events.length) return;

  const token = await getMicrosoftAccessToken();
  const batchRequests = events.map((event, index) => ({
    id: `event${index + 1}`,
    method: "POST",
    url: "/me/calendar/events",
    body: {
      subject: event.Event_Title,
      start: { dateTime: event.Start_DateTime, timeZone: "UTC" },
      end: { dateTime: event.End_DateTime, timeZone: "UTC" },
      location: { displayName: event.Venue || "" },
      attendees: event.Participants ? event.Participants.map(p => ({ emailAddress: { address: p.email }, type: "required" })) : []
    },
    headers: { "Content-Type": "application/json" }
  }));

  const batchSize = 20; // Microsoft allows up to 20 requests per batch
  for (let i = 0; i < batchRequests.length; i += batchSize) {
    const batchPayload = { requests: batchRequests.slice(i, i + batchSize) };

    await retryWithBackoff(async () => {
      await axios.post("https://graph.microsoft.com/v1.0/$batch", batchPayload, {
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }
      });
    });

    console.log(`✅ Batch ${i / batchSize + 1} synced successfully.`);
  }
}

// Function to Sync Zoho Events to Outlook
async function syncZohoToOutlook() {
  try {
    console.log("🔄 Fetching Zoho Events...");
    const zohoEvents = await fetchZohoEvents();
    if (!zohoEvents.length) {
      console.log("ℹ️ No new events found in Zoho.");
      return { success: true, message: "No new events found.", syncedEvents: [] };
    }

    console.log(`📅 Found ${zohoEvents.length} events. Syncing to Outlook...`);
    await createOutlookEvents(zohoEvents);

    console.log("✅ Sync completed successfully.");
    return { success: true, message: "Sync completed", syncedEvents: zohoEvents };
  } catch (error) {
    console.error("❌ Sync Failed:", error.response ? error.response.data : error.message);
    throw error;
  }
}

async function getOutlookEvents() {
  try {
    const accessToken = await getMicrosoftAccessToken();
    if (!accessToken) {
      console.error("❌ No valid Microsoft Access Token. Cannot fetch Outlook events.");
      return [];
    }

    // let events = [];
    let nextLink = "https://graph.microsoft.com/v1.0/me/events";
    const response = await axios.get(nextLink, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    // while (nextLink) {
    //     const response = await axios.get(nextLink, {
    //         headers: { Authorization: `Bearer ${accessToken}` }
    //     });

    //     events.push(...response.data.value);
    //     nextLink = response.data["@odata.nextLink"] || null;
    // }

    console.log(`✅ Fetched Outlook Events Successfully!`);
    return response.data.value;
  } catch (error) {
    console.error("❌ Error Fetching Outlook Events:", error.response ? error.response.data : error.message);
    return [];
  }
}


async function createZohoEvent(event) {
  if (!event.subject || !event.start?.dateTime || !event.end?.dateTime) {
    console.error("❌ Missing required fields:", event);
    return null;
  }

  const eventData = {
    data: [
      {
        Event_Title: event.subject,
        Start_DateTime: formatDate(event.start.dateTime),
        End_DateTime: formatDate(event.end.dateTime),
        Venue: event.location?.displayName || "N/A",
        Description: event.bodyPreview || "",
        Participants: event.attendees
          ? event.attendees.map((a) => ({
            type: "email",
            participant: a.emailAddress.address
          }))
          : [],
        send_notification: true
      }
    ]
  };

  try {
    const access_token = await getZohoAccessToken();
    if (!access_token) {
      console.error("❌ No Zoho Access Token Available.");
      return null;
    }

    const { ZOHO_API_BASE } = process.env;
    const response = await axios.post(`${ZOHO_API_BASE}/crm/v2/Events`, eventData, {
      headers: {
        "Authorization": `Zoho-oauthtoken ${access_token}`,
        "Content-Type": "application/json"
      }
    });

    console.log("✅ Zoho Event Created Successfully!");
    return eventData.data[0]; // Return created event details
  } catch (error) {
    console.error("❌ Zoho API Error:", error.response ? error.response.data : error.message);
    return null;
  }
}

// API Endpoint
app.get("/sync-and-return-events", async (req, res) => {
  try {

    console.log("🔄 Sync request received...");

    const events = await getOutlookEvents();
    let createdEvents = [];

    for (const event of events) {
      const syncedEvent = await createZohoEvent(event);
      if (syncedEvent) {
        createdEvents.push(syncedEvent);
      }
    }

    console.log("✅ Sync Complete!");

    res.json(createdEvents);
  } catch (error) {
    res.status(500).json({ success: false, error: error.response ? error.response.data : error.message });
  }
});



// Middleware to ensure Zoho token is available
const ensureZohoAccessToken = async (req, res, next) => {
  if (!tokenCache.zoho.accessToken) {
    try {
      await getZohoAccessToken();
    } catch (error) {
      return res.status(500).json({ error: "Failed to authenticate with Zoho CRM" });
    }
  }
  next();
};


// Create Zoho CRM Meeting
app.post("/create-meeting", ensureZohoAccessToken, async (req, res) => {
  const { subject, startDateTime, endDateTime, venue, description, participants, Who_Id } = req.body;

  // Check if required fields are provided
  if (!subject || !startDateTime || !endDateTime) {
    return res.status(400).json({ success: false, error: "Event Title, Start DateTime, and End DateTime are required." });
  }

  const event = {
    Who_Id: { id: Who_Id },  // Optional: Add the contact ID if available
    Event_Title: subject,  // Required field
    Start_DateTime: startDateTime,  // Required field
    End_DateTime: endDateTime,  // Required field
    Venue: venue || "N/A",  // Optional
    Description: description || "",  // Optional
    Participants: participants || []  // Optional
  };

  try {
    // Call the existing createZohoEvent function to create the event in Zoho CRM
    const result = await createZohoEvent(event);

    // Check Zoho response
    if (result && result.data && result.data[0].status === 'success') {
      res.json({ success: true, eventResponse: result.data[0] });
    } else {
      res.status(500).json({ success: false, error: "Failed to create the event in Zoho CRM." });
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message || "Failed to create the event." });
  }
});


// Fetch Customer Interaction Summary
app.get("/customer-summary", ensureZohoAccessToken, async (req, res) => {
  try {
    const response = await axios.get("https://www.zohoapis.in/crm/v2/Activities", {
      headers: { Authorization: `Zoho-oauthtoken ${tokenCache.zoho.accessToken}` }
    });

    const summary = response.data.data.map(activity => ({
      customer: activity.Who_Id?.name || "Unknown Customer",
      interaction: activity.Subject || "No Subject",
      date: activity.Created_Time || "No Date",
      status: activity.Status || "No Status"
    }));

    res.json({ summary });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post("/create-task", ensureZohoAccessToken, async (req, res) => {
  const { subject, dueDate, status, priority, description, Who_Id } = req.body;

  if (!subject || !dueDate) {
    return res.status(400).json({ success: false, error: "Subject and Due Date are required." });
  }

  const taskPayload = {
    Who_Id: { id: Who_Id },
    Subject: subject,
    Due_Date: dueDate,
    Status: status,
    Priority: priority,
    Description: description,
  };

  try {
    const response = await axios.post(`${process.env.ZOHO_API_BASE}/crm/v2/Tasks`, {
      data: [taskPayload]
    }, {
      headers: {
        Authorization: `Zoho-oauthtoken ${tokenCache.zoho.accessToken}`,
        "Content-Type": "application/json"
      }
    });

    res.json({ success: true, taskResponse: response.data });
  } catch (error) {
    console.error("❌ Task Creation Error:", error.response?.data || error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/zoho-leads', async (req, res) => {
  try {
    const accessToken = await getZohoAccessToken();
    const zohoResponse = await axios.get(`${process.env.ZOHO_API_BASE}/crm/v2/Leads`, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
      params: {
        sort_by: 'Created_Time',
        sort_order: 'desc',
        per_page: 5
      }
    });

    const leads = zohoResponse.data.data.map(lead => ({
      name: lead.Full_Name,
      email: lead.Email,
      company: lead.Company,
      phone: lead.Phone,
      status: lead.Lead_Status,
      created_at: lead.Created_Time
    }));

    res.json({ leads });
  } catch (err) {
    console.error("Zoho API Error (Leads):", err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to fetch leads from Zoho' });
  }
});

// Route: Fetch latest campaigns
app.get('/zoho-campaigns', async (req, res) => {
  try {
    const accessToken = await getZohoAccessToken();
    const zohoResponse = await axios.get(`${process.env.ZOHO_API_BASE}/crm/v2/Campaigns`, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
      params: {
        sort_by: 'Created_Time',
        sort_order: 'desc',
        per_page: 5
      }
    });

    const campaigns = zohoResponse.data.data.map(campaign => ({
      name: campaign.Campaign_Name,
      type: campaign.Type,
      status: campaign.Status,
      start_date: campaign.Start_Date,
      end_date: campaign.End_Date,
      expected_revenue: campaign.Expected_Revenue,
      actual_cost: campaign.Actual_Cost,
      budgeted_cost: campaign.Budgeted_Cost,
      created_at: campaign.Created_Time
    }));

    res.json({ campaigns });
  } catch (err) {
    console.error("Zoho API Error (Campaigns):", err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to fetch campaigns from Zoho' });
  }
});

app.get('/zoho-deals', async (req, res) => {
  try {
    const accessToken = await getZohoAccessToken();
    const response = await axios.get(`${process.env.ZOHO_API_BASE}/crm/v2/Deals`, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
      params: {
        sort_by: 'Created_Time',
        sort_order: 'desc',
        per_page: 5
      }
    });

    const deals = response.data.data.map(deal => ({
      id: deal.id,
      deal_name: deal.Deal_Name,
      description: deal.Description,
      stage: deal.Stage,
      type: deal.Type,
      probability: deal.Probability,
      amount: deal.Amount,
      expected_revenue: deal.Expected_Revenue,
      closing_date: deal.Closing_Date,
      created_at: deal.Created_Time,
      modified_at: deal.Modified_Time,
      overall_sales_duration: deal.Overall_Sales_Duration,
      sales_cycle_duration: deal.Sales_Cycle_Duration,
      lead_source: deal.Lead_Source,
      next_step: deal.Next_Step,
      campaign_source: deal.Campaign_Source,
      reason_for_loss: deal.Reason_For_Loss__s,
      owner: {
        id: deal.Owner?.id,
        name: deal.Owner?.name,
        email: deal.Owner?.email
      },
      account_name: deal.Account_Name?.name,
      account_id: deal.Account_Name?.id,
      created_by: {
        id: deal.Created_By?.id,
        name: deal.Created_By?.name,
        email: deal.Created_By?.email
      },
      modified_by: {
        id: deal.Modified_By?.id,
        name: deal.Modified_By?.name,
        email: deal.Modified_By?.email
      },
      approval_state: deal.$approval_state,
      is_editable: deal.$editable,
      layout_id: deal.$layout_id?.id,
      layout_name: deal.$layout_id?.name
    }));

    res.json({ deals });
  } catch (err) {
    console.error("Zoho API Error (Deals):", err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to fetch deals from Zoho' });
  }
});

app.get('/zoho-contacts', async (req, res) => {
  try {
    const accessToken = await getZohoAccessToken();
    const response = await axios.get(`${process.env.ZOHO_API_BASE}/crm/v2/Contacts`, {
      headers: { Authorization: `Zoho-oauthtoken ${accessToken}` },
      params: { sort_by: 'Created_Time', sort_order: 'desc', per_page: 10 }
    });

    const contacts = response.data.data.map(contact => ({
      id: contact.id,
      full_name: contact.Full_Name,
      email: contact.Email,
      phone: contact.Phone,
      mobile: contact.Mobile,
      title: contact.Title,
      department: contact.Department,
      mailing_address: {
        street: contact.Mailing_Street,
        city: contact.Mailing_City,
        state: contact.Mailing_State,
        zip: contact.Mailing_Zip,
        country: contact.Mailing_Country
      },
      created_at: contact.Created_Time
    }));

    res.json({ contacts });
  } catch (err) {
    console.error("Zoho API Error (Contacts):", err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to fetch contacts from Zoho' });
  }
});

app.put('/update-zoho-contacts/:id', async (req, res) => {
  try {
    const accessToken = await getZohoAccessToken();
    const recordId = req.params.id;
    const updateData = req.body;

    const body = {
      data: [updateData]
    };

    const response = await axios.put(`${process.env.ZOHO_API_BASE}/crm/v2/Contacts/${recordId}`, body, {
      headers: {
        Authorization: `Zoho-oauthtoken ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    res.json({ success: true, result: response.data });
  } catch (err) {
    console.error("Zoho API Error (Update Contact):", err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to update contact' });
  }
});

// Test API
app.get("/test-api", (req, res) => {
  res.json({ message: "Zoho CRM API is working!" });
});

// Root Page
app.get("/", (req, res) => {
  res.send(`<h1>Zoho CRM Task Automation</h1><p>Check API status <a href="/test-api">here</a>.</p>`);
});

// Start Server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
