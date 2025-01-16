from fastapi import APIRouter, Request
from typing import List

# 각 기능별 라우터 import (나중에 만들어질 것들)
# from app.api.v1.endpoints import users, auth, items ...

# 메인 라우터 생성
api_router = APIRouter()


# 예시 엔드포인트
@api_router.get("/hello")
async def hello_world():
    return {"message": "Hello World!"}


# 각 기능별 라우터 등록
# api_router.include_router(users.router, prefix="/users", tags=["users"]) # ex) app/api/v1/endpoints/users.py
# api_router.include_router(auth.router, prefix="/auth", tags=["auth"])
# api_router.include_router(items.router, prefix="/items", tags=["items"])


# Supabase 테스트용 엔드포인트 예시
@api_router.get("/test-db")
async def test_db(request: Request):
    try:
        supabase = request.state.supabase
        response = supabase.table("suppliers").select("*").limit(1).execute()
        return {"success": True, "suppliers": response.data}  # "suppliers":[{"id":2178,"supplier":"한국"}]일 때
        # return {
        #     "success": True,
        #     "suppliers_id": response.data[0]["id"],
        #     "suppliers_name": response.data[0]["supplier"],
        # }
    except Exception as e:
        return {"success": False, "error": str(e)}
