import time
from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware, RequestResponseEndpoint
from starlette.responses import Response
import logging

# 로깅 설정
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class LoggingMiddleware(BaseHTTPMiddleware):
    async def dispatch(
        self, request: Request, call_next: RequestResponseEndpoint
    ) -> Response:
        # 요청 시작 시간
        start_time = time.time()

        # 요청 정보 로깅
        logger.info(f"Request started: {request.method} {request.url}")

        # 요청 처리
        try:
            response = await call_next(request)

            # 처리 시간 계산
            process_time = time.time() - start_time

            # 응답 정보 로깅
            logger.info(
                f"Request completed: {request.method} {request.url} "
                f"- Status: {response.status_code} "
                f"- Processing Time: {process_time:.3f}s"
            )

            # 응답 헤더에 처리 시간 추가
            response.headers["X-Process-Time"] = str(process_time)

            return response

        except Exception as e:
            # 에러 로깅
            logger.error(
                f"Request failed: {request.method} {request.url} " f"- Error: {str(e)}"
            )
            raise
