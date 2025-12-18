# SYSMEX Workplace - Azure Functions Migration

## Migration Overview
Migrace z lokálního CLI nástroje na Azure Functions serverless architekturu.

### Branch Strategy
- **security_backup**: Functional backup před migrací (READ-ONLY)
- **azure-functions-migration**: Aktivní vývoj Azure Functions (CURRENT)
- **main**: Production branch (merge po dokončení migrace)

## Azure Functions Architecture

### Function Triggers
1. **HTTP Trigger** (`sap_process_http`)
   - Endpoint: `/api/sap_process`
   - Method: POST
   - Input: Multipart form-data s Excel souborem
   - Output: Zpracovaný Excel jako attachment
   - Use case: Manuální upload přes web UI nebo API

2. **Blob Trigger** (`sap_process_blob`)
   - Input: `input/` kontejner v Azure Blob Storage
   - Output: `output/` kontejner v Azure Blob Storage
   - Auto-spouštění při nahrání nového souboru
   - Use case: Automatizované batch zpracování

### Code Structure
```
function_app.py          # Main Azure Functions handlers
  ├── sap_process_http()           # HTTP trigger
  ├── sap_process_blob()           # Blob trigger
  └── process_sap_file_in_memory() # Core logic (sdílená)

utils.py                 # Shared utilities (z původního kódu)
sap_processor.py         # Core business logic (normalize_columns)
host.json                # Azure Functions runtime config
local.settings.json      # Local development settings (GITIGNORED)
.funcignore              # Files excluded from deployment
```

### Key Migration Changes
1. **File I/O → Memory Operations**
   - Původní: `process_file(input_path, output_path)` s disk I/O
   - Nový: `process_sap_file_in_memory(bytes) -> bytes`
   - Používá tempfile pro pandas/openpyxl kompatibilitu

2. **Storage**
   - Původní: Local filesystem `sklady_porovnani/input`, `/output`
   - Nový: Azure Blob Storage kontejnery nebo HTTP upload/download

3. **Configuration**
   - Původní: Hardcoded paths v `if __name__ == "__main__"`
   - Nový: Azure App Settings + local.settings.json

## Local Development

### Prerequisites
```bash
# Install Azure Functions Core Tools
# Linux: https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local
# nebo via package manager

# Install dependencies
pip install -r requirements.txt
```

### Running Locally
```bash
# Start Azure Functions runtime
func start

# HTTP endpoint bude dostupný na:
# http://localhost:7071/api/sap_process

# Test HTTP trigger:
curl -X POST http://localhost:7071/api/sap_process \
  -F "file=@sklady_porovnani/input/SAP.xlsx" \
  --output processed.xlsx
```

### Testing Blob Trigger Locally
- Requires Azurite (Azure Storage Emulator)
- Or connection string to real Azure Storage Account

## Deployment

### Azure Prerequisites
1. Azure Subscription
2. Resource Group
3. Function App (Python 3.12, Linux)
4. Storage Account (pro Blob trigger)

### Deploy Commands
```bash
# Login to Azure
az login

# Create Function App (example)
az functionapp create --resource-group <rg-name> \
  --consumption-plan-location westeurope \
  --runtime python --runtime-version 3.12 \
  --functions-version 4 \
  --name <function-app-name> \
  --storage-account <storage-account-name>

# Deploy
func azure functionapp publish <function-app-name>
```

### Configuration (Azure Portal)
Set Application Settings:
- `AzureWebJobsStorage`: Connection string k Storage Account
- `FUNCTIONS_WORKER_RUNTIME`: `python`

## Testing Strategy

### Pre-Migration Tests
- ✅ Original tests in `test_sap_processor.py` pass on `security_backup` branch

### Post-Migration Tests (TODO)
- [ ] Unit tests pro `process_sap_file_in_memory()`
- [ ] Integration test HTTP trigger (local)
- [ ] Integration test Blob trigger (local s Azurite)
- [ ] End-to-end test v Azure prostředí

## Backward Compatibility
- Original `sap_processor.py` zůstává funkční pro lokální použití
- Azure Functions jsou nová vrstva nad stejnou business logikou
- Sdílené funkce v `utils.py` a `normalize_columns()`

## Known Limitations
- HTTP trigger má timeout (default 5 min, max 10 min v host.json)
- Memory limit závislý na Azure Functions plan
- Pro velké soubory (>100MB) zvážit Durable Functions nebo jiný přístup
