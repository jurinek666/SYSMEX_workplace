# Azure Functions Migration - Session Notes
**Date:** December 18, 2025  
**Session Goal:** Prepare backup branch and initialize Azure Functions migration

## Session Summary

### 1. Backup Branch Preparation (`security_backup`)
**Goal:** Create stable backup before Azure Functions migration

**Actions Taken:**
- ✅ Identified and fixed Czech character bug in `sap_processor.py`
  - Changed: `"Nazev"` → `"Název"` (proper Czech ž character)
  - Location: Line 68 in rename_map
- ✅ Verified all tests pass successfully
  - Command: `python -m unittest test_sap_processor.py -v`
  - Result: OK (all tests passing)
- ✅ Committed fix to `security_backup` branch
  - Commit: "Fix: Oprava názvu sloupce Název (Czech character) - functional backup před migrací na Azure Functions"
- ✅ Pushed to remote: `origin/security_backup`

**Branch Status:**
- Branch: `security_backup` 
- Purpose: READ-ONLY functional backup
- State: Clean working tree, all tests passing
- Protection: Do not modify - reference point for rollback if needed

---

### 2. Azure Functions Migration Branch Setup
**Goal:** Create new development branch with Azure Functions scaffold

**Actions Taken:**
1. Created new branch from main:
   ```bash
   git checkout main
   git pull origin main
   git checkout -b azure-functions-migration
   ```

2. Created Azure Functions project structure:
   - ✅ `function_app.py` - Main Azure Functions handlers
   - ✅ `host.json` - Runtime configuration
   - ✅ `local.settings.json` - Local dev settings (gitignored)
   - ✅ `.funcignore` - Deployment exclusions
   - ✅ `.gitignore` - Added local.settings.json
   - ✅ `.github/copilot-instructions-azure.md` - Full migration docs

3. Updated dependencies:
   - Modified `requirements.txt` to include `azure-functions`

4. Committed initial scaffold:
   - Commit: "Azure Functions: Initial migration scaffold"
   - Files: 6 changed, 367 insertions(+)

**Branch Status:**
- Branch: `azure-functions-migration`
- Purpose: Active development for Azure migration
- Parent: `main` branch
- First Commit: 9619f3c

---

## Architecture Decisions

### Function Triggers Implemented
1. **HTTP Trigger** (`sap_process_http`)
   - Endpoint: POST `/api/sap_process`
   - Use Case: Manual file upload via API/web UI
   - Input: Multipart form-data with Excel file
   - Output: Processed Excel as download attachment

2. **Blob Trigger** (`sap_process_blob`)
   - Auto-executes on file upload to Azure Blob Storage
   - Input Container: `input/`
   - Output Container: `output/`
   - Use Case: Automated batch processing

### Key Migration Pattern
**Before (Local):**
```python
process_file(input_path, output_path)  # Disk I/O
```

**After (Azure Functions):**
```python
process_sap_file_in_memory(bytes) -> bytes  # Memory operations
```

**Rationale:**
- Azure Functions work best with in-memory operations
- Uses tempfile for pandas/openpyxl compatibility
- Automatic cleanup after processing
- Supports both HTTP response and Blob Storage output

---

## Next Steps (TODO)

### Immediate (Testing)
- [ ] Install Azure Functions Core Tools
- [ ] Test locally: `func start`
- [ ] Test HTTP endpoint with curl/Postman
- [ ] Verify output file matches original processor behavior

### Short-term (Development)
- [ ] Add unit tests for `process_sap_file_in_memory()`
- [ ] Migrate other processors:
  - [ ] `compare_processor.py` → Azure Function
  - [ ] `raben_processor.py` → Azure Function  
  - [ ] `merge_processor.py` → Azure Function
- [ ] Add error handling and logging improvements
- [ ] Consider adding input validation

### Medium-term (Deployment)
- [ ] Create Azure resources:
  - [ ] Resource Group
  - [ ] Function App (Python 3.12, Linux)
  - [ ] Storage Account (for Blob trigger)
  - [ ] Application Insights (optional, monitoring)
- [ ] Configure Azure App Settings
- [ ] Deploy: `func azure functionapp publish <app-name>`
- [ ] Test in Azure environment
- [ ] Set up CI/CD pipeline (GitHub Actions?)

### Long-term (Production)
- [ ] Merge `azure-functions-migration` → `main`
- [ ] Update documentation for production use
- [ ] Archive `security_backup` branch (keep for history)
- [ ] Monitor performance and costs in Azure
- [ ] Plan for scaling if needed

---

## Important Notes

### Preserved Original Functionality
- Original `sap_processor.py` remains functional for local CLI use
- Core business logic (`normalize_columns`, utils) is shared
- Backward compatible - both approaches work

### Configuration Management
- `local.settings.json` is gitignored (contains local configs)
- Azure production settings go in Azure Portal → App Settings
- Connection strings managed via Azure Key Vault (recommended)

### Known Limitations
- HTTP trigger timeout: 10 minutes max (configured in host.json)
- Memory limits depend on Azure Functions plan
- For files >100MB, consider Durable Functions or alternative approach

---

## Commands Reference

### Local Testing
```bash
# Start local Azure Functions runtime
func start

# Test HTTP endpoint
curl -X POST http://localhost:7071/api/sap_process \
  -F "file=@sklady_porovnani/input/SAP.xlsx" \
  --output processed.xlsx

# Run original tests (backup verification)
python -m unittest test_sap_processor.py -v
```

### Git Workflow
```bash
# Switch to backup (read-only)
git checkout security_backup

# Switch to development
git checkout azure-functions-migration

# View changes
git diff main..azure-functions-migration

# Push changes
git push origin azure-functions-migration
```

### Azure Deployment
```bash
# Login to Azure
az login

# Deploy to Azure Functions
func azure functionapp publish <function-app-name>

# View logs
func azure functionapp logstream <function-app-name>
```

---

## Session End State

**Branches:**
- `main` - Production baseline
- `security_backup` - Functional backup (commits: e1ca54e)
- `azure-functions-migration` - Active development (commits: 9619f3c)

**Working Directory:** Clean  
**Tests:** All passing on `security_backup`  
**Next Action:** Local testing of Azure Functions scaffold

---

## Questions to Consider

1. **Hosting Plan:** Consumption (pay-per-execution) vs Premium (always-on)?
2. **Authentication:** Current setup uses function-level auth keys. Need more security?
3. **Monitoring:** Should we set up Application Insights from the start?
4. **Storage:** Use same Storage Account for input/output or separate?
5. **Naming Convention:** How to name Azure resources? (e.g., `sysmex-sap-func-prod`)

---

**Session completed successfully. All changes committed and documented.**
