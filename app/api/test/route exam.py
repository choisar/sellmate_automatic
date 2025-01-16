from fastapi import APIRouter
from ....test.selenium_handler import process_stock_data

router = APIRouter()


@router.get("/process-stock")
async def process_stock():
    try:
        result_path = process_stock_data()
        return {"status": "success", "file_path": result_path}
    except Exception as e:
        return {"status": "error", "message": str(e)}
