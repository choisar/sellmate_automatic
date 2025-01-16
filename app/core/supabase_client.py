# core/database.py
from supabase import create_client, Client
from app.core.config import settings

# Supabase 연결 관리를 위한 클래스
class DatabaseClient:
    _client = None
    
    @classmethod
    def get_client(cls) -> Client:
        """
        Supabase 클라이언트를 가져옵니다.
        한 번 생성된 클라이언트는 재사용됩니다.
        """
        if cls._client is None:
            cls._client = cls._create_client()
        return cls._client
    
    @classmethod
    def _create_client(cls) -> Client:
        """새로운 Supabase 클라이언트를 만듭니다."""
        try:
            return create_client(
                settings.SUPABASE_URL,
                settings.SUPABASE_KEY
            )
        except Exception as e:
            raise Exception(f"Supabase 연결에 실패했습니다: {str(e)}")
    
    @classmethod
    def check_connection(cls) -> bool:
        """
        Supabase 연결이 정상적으로 작동하는지 확인합니다.
        """
        try:
            client = cls.get_client()
            # 간단한 쿼리로 연결 상태 확인
            client.table("suppliers").select("*").limit(1).execute()
            return True
        except Exception:
            return False
    
    @classmethod
    def reset_client(cls) -> None:
        """
        클라이언트를 초기화합니다. 
        연결 문제가 있을 때 사용할 수 있습니다.
        """
        cls._client = None

# 다른 파일에서 임포트해서 사용할 수 있는 클라이언트
supabase = DatabaseClient.get_client()