class HvdcError(Exception): ...
class ScanError(HvdcError): ...
class ParseError(HvdcError): ...
class PatternError(HvdcError): ...
class IoError(HvdcError): ...

# Outlook 관련 예외
class OutlookConnectionError(ScanError):
    """Outlook 연결 실패"""
    pass

class OutlookScanError(ScanError):
    """Outlook 스캔 오류"""
    pass

class OutlookPermissionError(ScanError):
    """Outlook 폴더 접근 권한 없음"""
    pass
