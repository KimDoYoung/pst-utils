# exceptions.py
class PstExtractError(Exception):
    """fund_mail 공통 최상위 예외"""

class TokenError(PstExtractError):
    """Graph API 토큰 발급 실패"""

class EmailFetchError(PstExtractError):
    """메일 목록/본문 조회 실패"""

class AttachFileFetchError(PstExtractError):
    """첨부파일 다운로드 실패"""

class DBCreateError(PstExtractError):
    """DB CREATE 실패"""

class DBWriteError(PstExtractError):
    """DB INSERT/UPDATE 실패"""

class DBQueryError(PstExtractError):
    """DB SELECT 실패"""

class SFTPUploadError(PstExtractError):
    """SFTP 업로드 실패"""
