# Financial Disclosure Generator

React + Express app that:
1. Stores xlsx templates in PostgreSQL
2. Accepts `{ Service, hs_object_id }` payloads
3. Fetches `consent_order_json` from HubSpot
4. Fills the correct template based on `Service`
5. POSTs the filled xlsx as binary to the n8n webhook

---

## Service → Template mapping

| Service       | Template                                          |
|---------------|---------------------------------------------------|
| `Assisted`    | Consent-Order---Summary-of-Financial-Disclosure   |
| `Negotiation` | Financial_Disclosure_and_Net_Effect_Table_Template|

---

## Setup

### 1. Prerequisites
- Node.js 18+
- PostgreSQL database
- HubSpot Private App token (scope: `crm.objects.deals.read`)

### 2. Install dependencies
```bash
npm run install:all
```

### 3. Configure environment
```bash
cp .env.example .env
# Edit .env with your actual values
```

Required variables:
```env
DATABASE_URL=postgresql://user:password@localhost:5432/financial_disclosure
HUBSPOT_API_KEY=your_hubspot_private_app_token
N8N_WEBHOOK_URL=https://n8n.amicablerd.uk/webhook/910c3d20-faee-41b9-96d0-e5716c011c5b
PORT=3001
```

### 4. Run (development)
```bash
npm run dev
# Server: http://localhost:3001
# Client: http://localhost:3000
```

### 5. Upload templates via the UI
1. Open http://localhost:3000
2. Go to **Templates** tab
3. Upload the Assisted .xlsx and Negotiation .xlsx files

---

## API

### Process endpoint (called by n8n or directly)
```
POST /api/process
Content-Type: application/json

{
  "Service": "Negotiation",
  "hs_object_id": 358206255303
}
```

Response:
```json
{
  "success": true,
  "message": "Negotiation template filled and sent to n8n",
  "filename": "Financial_Disclosure_Negotiation_Sherwood_2026-03-27.xlsx",
  "n8n_response": "..."
}
```

### Other endpoints
```
GET  /api/templates              — list templates
POST /api/templates              — upload template (multipart: service, name, file)
GET  /api/templates/:service/download  — download raw template
DEL  /api/templates/:service     — delete template
GET  /api/logs                   — last 50 processing records
GET  /api/health                 — health check
```

---

## Docker

```dockerfile
FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build
EXPOSE 3001
ENV NODE_ENV=production
CMD ["npm", "start"]
```

```yaml
# docker-compose.yml
services:
  app:
    build: .
    ports: ["3001:3001"]
    environment:
      DATABASE_URL: postgresql://postgres:password@db:5432/financial_disclosure
      HUBSPOT_API_KEY: ${HUBSPOT_API_KEY}
      NODE_ENV: production
    depends_on: [db]

  db:
    image: postgres:16
    environment:
      POSTGRES_DB: financial_disclosure
      POSTGRES_PASSWORD: password
    volumes:
      - pgdata:/var/lib/postgresql/data

volumes:
  pgdata:
```
