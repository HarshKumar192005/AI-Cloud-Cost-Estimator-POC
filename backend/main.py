# from typing import Optional

# from fastapi import FastAPI, UploadFile, File, HTTPException, Query
# from fastapi.middleware.cors import CORSMiddleware

# from azure_pricing import AzurePricingCalculator
# from oci_pricing import excel_to_requirements, fetch_products, run_oci_estimate


# from azure_pricing import AzurePricingCalculator
# from oci_pricing import excel_to_requirements, fetch_products, run_oci_estimate
# from aws_pricing import run_aws_estimate  # + new import

# app = FastAPI(title="Cloud Cost Estimator (Azure + OCI)")

# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],  # tighten this to your frontend's origin in production
#     allow_methods=["*"],
#     allow_headers=["*"],
# )


# def run_azure_estimate(file_bytes: bytes) -> dict:
#     """
#     Runs the (unchanged) AzurePricingCalculator against the same
#     'Category: ...' shaped Excel sheet used for OCI, so one uploaded
#     file can drive both engines.
#     """
#     reqs = excel_to_requirements(file_bytes)

#     if not reqs:
#         return {
#             "cloud": "Azure",
#             "total_monthly_cost": 0,
#             "currency": "USD",
#             "breakdown": [],
#             "warnings": ["No requirements found in the uploaded sheet"],
#         }

#     calculator = AzurePricingCalculator()
#     result = calculator.calculate(reqs)
#     result["warnings"] = []
#     return result


# @app.get("/api/health")
# def health():
#     return {"status": "ok"}


# @app.post("/api/refresh-products")
# def refresh_products():
#     """Refreshes the cached OCI product price list."""
#     products = fetch_products(force=True)
#     return {"count": len(products)}





# @app.post("/api/estimate")
# async def estimate(
#     file: UploadFile = File(...),
#     cloud: Optional[str] = Query(
#         "all",
#         description="Which engine(s) to run: 'azure', 'oci', 'aws', or 'all' (default).",
#     ),
# ):
#     if not file.filename.lower().endswith((".xlsx", ".xlsm")):
#         raise HTTPException(400, "Please upload an .xlsx file")

#     cloud = (cloud or "all").lower().strip()
#     valid = ("azure", "oci", "aws", "all", "both")  # 'both' kept for backward compat = azure+oci
#     if cloud not in valid:
#         raise HTTPException(400, "cloud must be one of: azure, oci, aws, all")

#     file_bytes = await file.read()
#     response: dict = {"cloud": cloud}

#     if cloud in ("azure", "both", "all"):
#         try:
#             response["azure"] = run_azure_estimate(file_bytes)
#         except Exception as e:
#             response["azure"] = {"error": str(e)}

#     if cloud in ("oci", "both", "all"):
#         try:
#             response["oci"] = run_oci_estimate(file_bytes)
#         except Exception as e:
#             response["oci"] = {"error": str(e)}

#     if cloud in ("aws", "all"):
#         try:
#             response["aws"] = run_aws_estimate(file_bytes)
#         except Exception as e:
#             response["aws"] = {"error": str(e)}

#     return response
import io
import json
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from azure_pricing import AzurePricingCalculator
from oci_pricing import excel_to_requirements, fetch_products, run_oci_estimate

app = FastAPI(title="Cloud Cost Estimator (Azure + OCI)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten this to your frontend's origin in production
    allow_methods=["*"],
    allow_headers=["*"],
)


def run_azure_estimate(file_bytes: bytes) -> dict:
    """
    Runs the (unchanged) AzurePricingCalculator against the same
    'Category: ...' shaped Excel sheet used for OCI, so one uploaded
    file can drive both engines.
    """
    reqs = excel_to_requirements(file_bytes)

    if not reqs:
        return {
            "cloud": "Azure",
            "total_monthly_cost": 0,
            "currency": "USD",
            "breakdown": [],
            "warnings": ["No requirements found in the uploaded sheet"],
        }

    calculator = AzurePricingCalculator()
    result = calculator.calculate(reqs)
    result["warnings"] = []
    return result


def build_workbook(response: dict) -> io.BytesIO:
    """One workbook: Summary sheet + one sheet per cloud (Azure / OCI)."""
    wb = Workbook()
    wb.remove(wb.active)
    summary = wb.create_sheet("Summary")
    summary.append(["Cloud", "Total (USD)", "Status"])

    for key, title in (("azure", "Azure"), ("oci", "OCI")):
        result = response.get(key)
        if result is None:
            continue
        ws = wb.create_sheet(title)

        if "error" in result:
            ws.append(["Error", result["error"]])
            summary.append([title, None, "Error"])
            continue

        # OCI uses "report"; Azure uses "breakdown"
        rows = result.get("report") or result.get("breakdown") or []
        cols = []
        for r in rows:
            for k in r:
                if k not in cols:
                    cols.append(k)

        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True)
        for r in rows:
            ws.append([
                json.dumps(r.get(c)) if isinstance(r.get(c), (dict, list)) else r.get(c)
                for c in cols
            ])

        total = result.get("total", result.get("total_monthly_cost", 0))
        cost_idx = cols.index("cost") + 1 if "cost" in cols else max(len(cols), 2)
        ws.append([])
        ws.cell(row=ws.max_row + 1, column=1, value="TOTAL (USD)").font = Font(bold=True)
        tcell = ws.cell(row=ws.max_row, column=cost_idx, value=total)
        tcell.font = Font(bold=True)

        for w in result.get("warnings", []):
            ws.append(["Warning", w])

        for i, col in enumerate(ws.columns, start=1):
            longest = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            ws.column_dimensions[get_column_letter(i)].width = min(longest + 2, 60)

        summary.append([title, total, "OK"])

    for c in summary[1]:
        c.font = Font(bold=True)
    summary.column_dimensions["A"].width = 14
    summary.column_dimensions["B"].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/refresh-products")
def refresh_products():
    """Refreshes the cached OCI product price list."""
    products = fetch_products(force=True)
    return {"count": len(products)}


@app.post("/api/estimate")
async def estimate(
    file: UploadFile = File(...),
    cloud: Optional[str] = Query(
        "all",
        description="Which engine(s) to run: 'azure', 'oci', or 'all' (default).",
    ),
):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Please upload an .xlsx file")

    cloud = (cloud or "all").lower().strip()
    valid = ("azure", "oci", "all", "both")  # 'both' kept for backward compat = azure+oci
    if cloud not in valid:
        raise HTTPException(400, "cloud must be one of: azure, oci, all")

    file_bytes = await file.read()
    response: dict = {"cloud": cloud}

    if cloud in ("azure", "both", "all"):
        try:
            response["azure"] = run_azure_estimate(file_bytes)
        except Exception as e:
            response["azure"] = {"error": str(e)}

    if cloud in ("oci", "both", "all"):
        try:
            response["oci"] = run_oci_estimate(file_bytes)
        except Exception as e:
            response["oci"] = {"error": str(e)}

    return response


@app.post("/api/estimate/xlsx")
async def estimate_xlsx(
    file: UploadFile = File(...),
    cloud: Optional[str] = Query("all", description="'azure', 'oci', or 'all'"),
):
    """Same as /api/estimate, but returns one .xlsx: Summary + Azure and OCI sheets."""
    response = await estimate(file=file, cloud=cloud)
    buf = build_workbook(response)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="cloud_cost_estimate.xlsx"'},
    )