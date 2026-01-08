# Microsoft Integration Guide for BSA-CRM Performance Tracking

This comprehensive guide explains how to automatically update your OneNote pages and integrate Microsoft tools (OneNote, To Do, Teams, Outlook, Planner) with your BSA-CRM performance tracking system.

---

## Table of Contents

1. [Integration Overview](#integration-overview)
2. [Power Automate Workflows](#power-automate-workflows)
3. [Microsoft Graph API](#microsoft-graph-api)
4. [Microsoft Loop Components](#microsoft-loop-components)
5. [Ready-to-Use Workflow Templates](#ready-to-use-workflow-templates)
6. [Implementation Guide](#implementation-guide)
7. [Best Practices](#best-practices)

---

## Integration Overview

There are three primary methods to automatically update OneNote and integrate Microsoft tools:

| Method | Complexity | Best For | Requirements |
|--------|------------|----------|--------------|
| **Power Automate** | Low | No-code automation, triggers, scheduled updates | Microsoft 365 subscription |
| **Microsoft Graph API** | High | Custom applications, programmatic control | Developer skills, Azure AD app |
| **Microsoft Loop** | Low | Real-time collaboration, live components | Microsoft 365 subscription |

---

## Power Automate Workflows

Power Automate is the recommended approach for most users. It provides no-code automation with pre-built connectors for all Microsoft 365 services.

### Available OneNote Actions

| Action | Description | Use Case |
|--------|-------------|----------|
| Create page in a section | Creates a new OneNote page | Weekly reports, meeting notes |
| Update page content | Modifies existing page HTML | Auto-update dashboards |
| Get page content | Retrieves page HTML | Extract data for processing |
| Delete a page | Removes a page | Archive old content |
| Create section in notebook | Adds new section | Organize by project/month |

### Available OneNote Triggers

| Trigger | Description | Use Case |
|---------|-------------|----------|
| When a new page is created | Fires when page added | Notify team, log activity |
| When a new section is created | Fires when section added | Update index, notify |
| When a new section group is created | Fires when group added | Organizational updates |

### Key Power Automate Connectors

| Connector | Integration Capabilities |
|-----------|-------------------------|
| **OneNote (Business)** | Create/update pages, sections, get content |
| **Microsoft To Do** | Create tasks, update status, get lists |
| **Outlook** | Email triggers, calendar events, send notifications |
| **Teams** | Post messages, meeting triggers, channel updates |
| **Planner** | Create/update tasks, buckets, plans |
| **Excel Online** | Read/write data, trigger on changes |
| **SharePoint** | Document triggers, list updates |

---

## Recommended Workflows for BSA-CRM Role

### Workflow 1: Auto-Create Weekly Performance Summary

**Trigger:** Scheduled (Every Friday at 4 PM)

**Actions:**
1. Get data from Excel tracker (BSA_CRM_OPTIMIZED_Dashboard.xlsx)
2. Format HTML content with KPIs
3. Create new OneNote page in "Weekly Pulse" section
4. Send summary email to manager

**Power Automate Steps:**
```
Trigger: Recurrence (Weekly, Friday, 4:00 PM)
   ↓
Action: Excel Online - List rows present in a table
   ↓
Action: Compose - Format HTML with data
   ↓
Action: OneNote - Create page in a section
   ↓
Action: Outlook - Send an email
```

### Workflow 2: Meeting Notes to OneNote

**Trigger:** When Teams meeting ends

**Actions:**
1. Get meeting details and transcript
2. Create OneNote page with meeting notes template
3. Extract action items
4. Create To Do tasks for each action item

**Power Automate Steps:**
```
Trigger: Teams - When a meeting ends
   ↓
Action: Teams - Get meeting details
   ↓
Action: OneNote - Create page in a section
   ↓
Action: Microsoft To Do - Add a to-do (for each action item)
```

### Workflow 3: Email to Task Automation

**Trigger:** When flagged email arrives

**Actions:**
1. Extract email subject and body
2. Create To Do task with deadline
3. Log to OneNote activity tracker
4. Send confirmation

**Power Automate Steps:**
```
Trigger: Outlook - When an email is flagged
   ↓
Action: Microsoft To Do - Add a to-do
   ↓
Action: OneNote - Update page content (Activity Log)
   ↓
Action: Outlook - Send an email (confirmation)
```

### Workflow 4: Daily Activity Log Sync

**Trigger:** Scheduled (Daily at 6 PM)

**Actions:**
1. Get today's activities from Excel
2. Format as HTML list
3. Append to OneNote Daily Activity Log page
4. Update dashboard metrics

**Power Automate Steps:**
```
Trigger: Recurrence (Daily, 6:00 PM)
   ↓
Action: Excel Online - List rows (filter by today)
   ↓
Action: Compose - Format HTML
   ↓
Action: OneNote - Update page content
```

### Workflow 5: Sprint Status to Teams

**Trigger:** When OneNote page updated (Sprint Tracker)

**Actions:**
1. Get updated page content
2. Extract sprint status
3. Post update to Teams channel
4. Update Planner tasks

**Power Automate Steps:**
```
Trigger: OneNote - When a new page is created
   ↓
Action: OneNote - Get page content
   ↓
Action: Teams - Post message in a chat or channel
   ↓
Action: Planner - Update a task
```

---

## Microsoft Graph API

For advanced users who need programmatic control, Microsoft Graph API provides full access to OneNote.

### API Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/me/onenote/notebooks` | GET | List all notebooks |
| `/me/onenote/sections` | GET | List all sections |
| `/me/onenote/pages` | GET/POST | List or create pages |
| `/me/onenote/pages/{id}/content` | GET/PATCH | Get or update page content |

### Python Example: Create OneNote Page

```python
import requests

# Authentication (requires Azure AD app registration)
access_token = "YOUR_ACCESS_TOKEN"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/xhtml+xml"
}

# HTML content for the page
html_content = """
<!DOCTYPE html>
<html>
<head>
    <title>Weekly Performance Summary</title>
</head>
<body>
    <h1>Week 2 Performance Summary</h1>
    <p>Tasks Completed: 12</p>
    <p>On-Time Rate: 95%</p>
    <table>
        <tr><th>Category</th><th>Count</th></tr>
        <tr><td>CRM Analysis</td><td>5</td></tr>
        <tr><td>System Admin</td><td>3</td></tr>
        <tr><td>Data Integration</td><td>4</td></tr>
    </table>
</body>
</html>
"""

# Create page in specific section
section_id = "YOUR_SECTION_ID"
url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"

response = requests.post(url, headers=headers, data=html_content)
print(response.json())
```

### Python Example: Update OneNote Page

```python
import requests
import json

access_token = "YOUR_ACCESS_TOKEN"
page_id = "YOUR_PAGE_ID"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# PATCH request to update page content
update_data = [
    {
        "target": "body",
        "action": "append",
        "content": "<p>New update added: Task completed at 3:00 PM</p>"
    }
]

url = f"https://graph.microsoft.com/v1.0/me/onenote/pages/{page_id}/content"
response = requests.patch(url, headers=headers, data=json.dumps(update_data))
print(response.status_code)
```

### Authentication Setup

1. Register an application in Azure AD
2. Configure API permissions (Notes.ReadWrite, Notes.Create)
3. Generate client secret or certificate
4. Use OAuth 2.0 flow to obtain access token

**Important:** As of March 31, 2025, Microsoft Graph OneNote API requires delegated authentication (app-only authentication is no longer supported).

---

## Microsoft Loop Components

Microsoft Loop provides real-time collaborative components that sync across Microsoft 365 apps.

### Benefits for BSA-CRM Tracking

| Feature | Benefit |
|---------|---------|
| Live sync | Changes appear instantly across all apps |
| Embed anywhere | Use in OneNote, Teams, Outlook, Word |
| Collaborative | Multiple people can edit simultaneously |
| Portable | Components work across Microsoft 365 |

### How to Use Loop in OneNote

1. Open OneNote (Windows or Web)
2. Click **Insert** → **Loop Component**
3. Choose component type (Task List, Table, Checklist)
4. Edit the component - changes sync everywhere

### Loop Component Types for Performance Tracking

| Component | Use Case |
|-----------|----------|
| **Task List** | Track action items with assignees and due dates |
| **Table** | Create live KPI dashboards |
| **Checklist** | Daily/weekly checklists |
| **Bulleted List** | Meeting notes, updates |
| **Numbered List** | Prioritized tasks |
| **Paragraph** | Shared notes and updates |

### Sharing Loop Components

1. Create a Loop component in OneNote
2. Copy the component
3. Paste into Teams chat, Outlook email, or Word document
4. All instances stay synchronized

---

## Ready-to-Use Workflow Templates

### Template 1: Weekly Report Generator

**Name:** BSA-CRM Weekly Report to OneNote

**Description:** Automatically generates a weekly performance report in OneNote every Friday.

**Setup Steps:**
1. Go to [Power Automate](https://flow.microsoft.com)
2. Click **Create** → **Scheduled cloud flow**
3. Set recurrence to Weekly, Friday, 4:00 PM
4. Add action: Excel Online → List rows present in a table
5. Add action: Compose → Format HTML report
6. Add action: OneNote → Create page in a section
7. Save and test

### Template 2: Email to Task Converter

**Name:** Flagged Email to To Do Task

**Description:** Creates a To Do task when you flag an email in Outlook.

**Setup Steps:**
1. Go to Power Automate
2. Search for template "Create a task in To Do when an email is flagged"
3. Connect your accounts
4. Customize task list and due date settings
5. Save and enable

### Template 3: Teams Meeting Notes

**Name:** Auto-Create Meeting Notes Page

**Description:** Creates a OneNote page with meeting template when a Teams meeting starts.

**Setup Steps:**
1. Go to Power Automate
2. Create new flow with trigger "When a Teams meeting starts"
3. Add action: OneNote → Create page in a section
4. Use meeting subject as page title
5. Include attendees and agenda in page content
6. Save and test

### Template 4: Daily Activity Sync

**Name:** Sync Excel Activities to OneNote

**Description:** Appends today's activities from Excel to OneNote daily log.

**Setup Steps:**
1. Create scheduled flow (Daily, 6:00 PM)
2. Get rows from Excel where date = today
3. Format as HTML list
4. Update OneNote page content (append)
5. Save and enable

---

## Implementation Guide

### Step 1: Set Up Power Automate

1. Go to [flow.microsoft.com](https://flow.microsoft.com)
2. Sign in with your Microsoft 365 work account
3. Click **Create** to start building flows

### Step 2: Connect Your Services

1. When adding actions, you'll be prompted to sign in
2. Authorize each connector (OneNote, To Do, Outlook, Teams)
3. Connections are saved for future use

### Step 3: Create Your First Flow

**Recommended First Flow:** Email to Task

1. Click **Create** → **Automated cloud flow**
2. Name: "Flagged Email to Task"
3. Trigger: Outlook → When an email is flagged
4. Action: Microsoft To Do → Add a to-do
5. Map email subject to task title
6. Save and test by flagging an email

### Step 4: Build OneNote Integration

1. Create a flow with your preferred trigger
2. Add action: OneNote (Business) → Create page in a section
3. Select your notebook and section
4. Format page content as HTML
5. Test and refine

### Step 5: Monitor and Optimize

1. Check flow run history for errors
2. Use the "Test" button to debug
3. Add error handling with "Configure run after"
4. Set up notifications for failed runs

---

## Best Practices

### For Power Automate

| Practice | Description |
|----------|-------------|
| Use descriptive names | Name flows clearly (e.g., "Weekly Report to OneNote") |
| Add error handling | Configure "run after" for failed actions |
| Test incrementally | Test each action before adding more |
| Use variables | Store values for reuse across actions |
| Document your flows | Add comments and descriptions |

### For OneNote Integration

| Practice | Description |
|----------|-------------|
| Use consistent HTML | Format pages consistently for easy parsing |
| Create dedicated sections | Separate automated content from manual notes |
| Use templates | Create page templates for consistent structure |
| Archive old content | Move completed items to archive sections |

### For Task Management

| Practice | Description |
|----------|-------------|
| Use priority levels | Set High/Medium/Low for all tasks |
| Set due dates | Always include deadlines |
| Use consistent naming | Follow naming conventions for tasks |
| Review regularly | Check automated tasks daily |

---

## Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| Flow not triggering | Check trigger conditions, verify connections |
| OneNote page not created | Verify section ID, check permissions |
| HTML not rendering | Validate HTML syntax, use simple formatting |
| Connection expired | Re-authorize the connector |
| Throttling errors | Add delays between actions, reduce frequency |

### Getting Help

- [Power Automate Documentation](https://docs.microsoft.com/en-us/power-automate/)
- [Microsoft Graph OneNote API](https://docs.microsoft.com/en-us/graph/integrate-with-onenote)
- [Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Community/ct-p/MPACommunity)

---

## Quick Reference Card

### Power Automate Shortcuts

| Action | How To |
|--------|--------|
| Create flow | flow.microsoft.com → Create |
| Test flow | Open flow → Test → Manually |
| View history | Open flow → Run history |
| Disable flow | Open flow → Turn off |

### OneNote HTML Tags

| Tag | Purpose |
|-----|---------|
| `<title>` | Page title |
| `<h1>` to `<h6>` | Headings |
| `<p>` | Paragraphs |
| `<table>` | Tables |
| `<ul>`, `<ol>` | Lists |
| `<img>` | Images |
| `<a>` | Links |

### Microsoft To Do Actions

| Action | Description |
|--------|-------------|
| Add a to-do | Create new task |
| Update a to-do | Modify existing task |
| Complete a to-do | Mark task as done |
| Get to-dos | List tasks from a list |

---

*Last Updated: January 2026*
