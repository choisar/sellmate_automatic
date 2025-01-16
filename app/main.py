from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.core.config import settings
from app.api.test.route import api_router
from app.middleware.logging import LoggingMiddleware
from app.core.supabase_client import supabase

# FastAPI 애플리케이션 인스턴스 생성
app = FastAPI(
    title=settings.PROJECT_NAME,
    description=settings.PROJECT_DESCRIPTION,
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS 미들웨어 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 커스텀 로깅 미들웨어 추가
app.add_middleware(LoggingMiddleware)

# API 라우터 등록
app.include_router(api_router, prefix="/api/test")

# Supabase 클라이언트를 의존성으로 주입
@app.middleware("http")
async def add_supabase_client(request, call_next):
    request.state.supabase = supabase
    response = await call_next(request)
    return response

# 헬스 체크 엔드포인트
@app.get("/health")
async def health_check():
    return {"status": "healthy"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", 
                host=settings.SERVER_HOST,
                port=settings.SERVER_PORT,
                reload=settings.DEBUG_MODE)