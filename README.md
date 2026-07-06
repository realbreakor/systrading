# systrading

LS증권 xingAPI를 이용한 파이썬 자동매매 라이브러리입니다.
주간/야간 선물·옵션의 현재가 조회, 주문, 잔고 조회 등을 지원합니다.

## 환경 요구사항

- Windows (xingAPI는 Windows 전용 COM 라이브러리입니다)
- Python 3.8 이상
- LS증권 계좌 및 xingAPI 설치 (`C:\LS_SEC\xingAPI`)

## 설치

```bash
git clone https://github.com/realbreakor/systrading.git
cd systrading
pip install -r requirements.txt
```

## 인증 정보 설정

`.env.example`을 복사해서 `.env` 파일을 만들고, 본인의 정보를 입력합니다.

```bash
cp .env.example .env
```

`.env` 파일:

```
APP_KEY=발급받은_앱키
SECRET_KEY=발급받은_시크릿키
ID=LS증권_로그인_아이디
ID_PSWD=LS증권_로그인_비밀번호
PSWD=계좌_비밀번호
CERT=공인인증서_비밀번호
```

> `.env` 파일은 절대 깃에 올리지 마세요. `.gitignore`에 이미 등록되어 있습니다.

## 사용 예시

모든 TR 함수는 `async` 함수입니다. `asyncio` 이벤트 루프 안에서 `await`로 호출해야 합니다.

### 로그인

```python
import asyncio
from ls_api import xingLogin, xingAccount

async def main():
    await xingLogin()
    acnt = await xingAccount("1234")  # 계좌번호 뒤 4자리

asyncio.run(main())
```

### 현재가 조회

```python
from ls_api import xingCode, xingCurrentPrice

async def main():
    code = (await xingCode())[0]
    price, open_price, last_month = await xingCurrentPrice(code)
    print(f"현재가: {price}, 시가: {open_price}")
```

### 시장가 주문

```python
from ls_api import xingFutOrder

async def main():
    await xingFutOrder(acnt, "2", 1, code)  # 매수
    await xingFutOrder(acnt, "1", 1, code)  # 매도
```

### 잔고 조회

```python
from ls_api import xingJango

async def main():
    avail = await xingJango(acnt)
    print(f"주문가능금액: {avail}")
```

## 주요 함수 목록

| 함수 | 설명 |
|---|---|
| `xingLogin()` | LS증권 서버 로그인 |
| `xingAccount(digit)` | 계좌번호 조회 (뒤 4자리로 검색) |
| `xingCurrentPrice(code)` | 현재가, 시가, 최종결제일 조회 |
| `xingBidAsk(code)` | 호가 5단계 조회 |
| `xingFutOrder(acnt, dir, qty, code)` | 선물 주문 (dir: `"1"`=매도, `"2"`=매수) |
| `xingCancelOrder(acnt, code, ordNo, qty)` | 주문 취소 |
| `xingJango(acnt)` | 주문가능금액 조회 |
| `xingCheckAccount(acnt, code)` | 보유 잔고 및 진입방향 조회 |
| `xingCheckOrder(acnt, code)` | 미체결 주문 조회 |
| `xingCode()` | 코스피200 선물 근월물 코드 조회 |
| `xingMiniCode()` | 미니 선물 근월물 코드 조회 |
| `xingUsdCode()` | 달러 선물 코드 조회 |

## 실시간 데이터 클래스

| 클래스 | 설명 |
|---|---|
| `XReal_JIF` | 장 상태 실시간 |
| `XReal_FX9` | 선물 가격제한폭 실시간 |
| `XReal_FC9` | 선물 현재가 실시간 |
| `XReal_OC0` | 옵션 현재가 실시간 |
| `XReal_C01` | 주간 체결 실시간 |
| `XReal_FH9` | 선물 호가 실시간 |
| `XReal_YF9` | 선물 예상체결가 실시간 |
| `XReal_YOC` | 옵션 예상체결가 실시간 |
| `XReal_DC0` | 야간 선물 현재가 실시간 |
| `XReal_C02` | 야간 체결 실시간 |
| `XReal_DH0` | 야간 선물 호가 실시간 |
