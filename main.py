from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

from schedule_reconciliation import schedule_reconciliation

# FastAPI app initializaiton
app = FastAPI(
    title="Schedule Reconciliation API",
    description="API for comparing schedule and workload Excel files",
    version="1.0.0"
)

# GET-request, returning API status
@app.get("/")
async def root():
    return {
        "message": "Schedule Reconciliation API is running"
    }

# POST-Request for full schedule reconciliation
@app.post("/reconcile")
async def reconcile(
    schedule_file: UploadFile = File(...),
    workload_file: UploadFile = File(...)
):
    # Simple file-type validation mechanism
    if not schedule_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Schedule file must be .xls or .xlsx"
        )

    if not workload_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Workload file must be .xls or .xlsx"
        )

    try:
        # Reset of file pointers to ensure reading files from the beginning
        schedule_file.file.seek(0)
        workload_file.file.seek(0)

        # Initialization of reconciler class with provided files
        reconciler = schedule_reconciliation(
            workload_file=workload_file.file,
            schedule_file=schedule_file.file
        )
        # Call of complete reconciliation function
        result = reconciler.full_check()

        return JSONResponse(content=result)

    # Exception in case of an internal processing error
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal processing error: {str(e)}"
        )

# POST Request featuring only mismatches
@app.post("/reconcile_mismatch")
async def reconcile_mismatch(
    schedule_file: UploadFile = File(...), 
    workload_file: UploadFile = File(...)
):
    # Simple file-type validation mechanism
    if not schedule_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Schedule file must be .xls or .xlsx"
        )

    if not workload_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Workload file must be .xls or .xlsx"
        )

    try:
        # Reset of file pointers to ensure reading files from the beginning
        schedule_file.file.seek(0)
        workload_file.file.seek(0)

        # Initialization of reconciler class with provided files
        reconciler = schedule_reconciliation(
            workload_file=workload_file.file,
            schedule_file=schedule_file.file
        )

        # Call of reconciliation function, returning only the list of classes with mismatched professors
        result = reconciler.mismatch_check()

        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal processing error: {str(e)}"
        )

# POST Request getting no-prof errors
@app.post("/reconcile_no_prof")
async def reconcile_no_prof(
    schedule_file: UploadFile = File(...), 
    workload_file: UploadFile = File(...)
):
    # Simple file-type validation mechanism
    if not schedule_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Schedule file must be .xls or .xlsx"
        )

    if not workload_file.filename.endswith((".xls", ".xlsx")):
        raise HTTPException(
            status_code=400,
            detail="Workload file must be .xls or .xlsx"
        )

    try:
        # Reset of file pointers to ensure reading files from the beginning
        schedule_file.file.seek(0)
        workload_file.file.seek(0)

        # Initialization of reconciler class with provided files
        reconciler = schedule_reconciliation(
            workload_file=workload_file.file,
            schedule_file=schedule_file.file
        )
        # Call of reconciliation function, returning only the list of classes missing a professor
        result = reconciler.no_prof_check()

        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Internal processing error: {str(e)}"
        )