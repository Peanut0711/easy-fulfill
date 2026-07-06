# 배송추적 자동 새로고침 (백그라운드 · 다중 PC 소프트 락)

작성일: 2026-07-06

배송추적 탭의 새로고침(우체국 종추적조회)이 그동안 **수동 버튼**으로만 돌았는데, 이를
**N분마다 백그라운드로 자동** 수행하도록 추가했다. 5인 공유 환경에서 공통 regkey에 부하가
몰리지 않도록 **공유 시트 타임스탬프로 다중 PC 중복 조회를 막는다**(소프트 락).

구현: `easy-fulfill.py`
- 워커 `run_tracking_refresh_worker(regkey, progress_cb, auto, interval_min)`
- 타이머/핸들러 `_tracking_auto_refresh_timer`, `_on_tracking_auto_refresh_tick`,
  `_on_tracking_auto_refresh_finished`, `_reconfigure_auto_refresh_timer`,
  `_auto_refresh_interval_min`
- 설정 UI `_dlg_auto_refresh`(설정 팝업), `_on_auto_refresh_interval_changed`,
  `_sync_auto_refresh_spinbox`
- 설정 헬퍼 `_write_config_values(ws, updates)`, `_parse_ts(s)`

## 동작 개요

- 앱 시작 시 타이머를 간격만큼 걸어 둔다. **첫 조회는 간격 경과 후**라 시작 직후엔 안 돈다.
- 타이머가 발화하면(`_on_tracking_auto_refresh_tick`) 미완료 송장을 우체국으로 백그라운드
  조회한다. **탭을 보고 있지 않아도** 돈다(백그라운드가 목적).
- 결과는 **팝업 없이** 요약 라벨만 갱신 + 표 재로딩 + 위험건 슬랙 다이제스트 검토
  (기존 하루 1통 · 시트 중복방지 재사용). **실패는 콘솔 로그만** 남기고 다음 주기에 재시도 →
  작업 흐름을 끊지 않는다.
- 수동 새로고침과 **같은 단일 실행 슬롯**(`_tracking_refresh_thread`)을 공유한다. 하나가
  도는 중이면 다른 쪽은 스킵 → 겹침 없음.

## 간격 설정 (전원 공유)

- 공유 「설정」 탭 키 **`tracking_auto_refresh_min`** (분). 기본 **60**, **`0`이면 끔**(수동만).
- 설정 팝업에 「배송추적 자동 새로고침」 간격 스핀박스(0~720분, `0`은 "끔" 표시). 변경 시
  로컬 저장 + 공유 시트 반영 + 타이머 재설정, **전원에게 적용**.
- 공유 설정 수신(`_on_tracking_config_read_finished`) 시 간격을 반영하고 타이머를 재구성한다.
- 워커는 조회 직전 설정 탭의 최신 간격을 다시 읽어 **caller가 넘긴 값보다 우선** 적용한다
  (운영 중 조정 즉시 반영).

## 다중 PC 중복 조회 방지 — 소프트 락

**문제**: 5대가 각자 1시간마다 자동 조회하면 공통 regkey에 시간당 5회가 몰려
우체국 부하 차단(`ERR-131`)이 더 자주 뜬다.

**해결**: 공유 「설정」 탭 키 **`tracking_last_auto_refresh`**(`YYYY-MM-DD HH:MM:SS`)를 락으로 쓴다.
자동 조회(`auto=True`)일 때 워커는:

1. 설정 탭에서 마지막 자동조회 시각을 읽는다.
2. **간격 안이면(`now - last < interval`) 우체국 호출 없이 스킵** → `{ok, skipped_recent}`.
3. 통과하면 **즉시 지금 시각을 기록해 슬롯을 선점**한 뒤 조회를 시작한다.

효과: 여러 대가 켜져 있어도 **간격이 지난 뒤 먼저 도는 1대만** 실제 API를 호출하고 나머지는
건너뛴다 → 공통키 부하 5회→1회. 조회가 실패하거나 `ERR-131`로 중단돼도 시각은 기록되므로
부하 상황에서 여러 대가 재차 몰리지 않는다.

### 한계 (의도한 트레이드오프)

- **완전한 분산 락이 아니다.** 두 PC의 타이머가 거의 동시에 발화하면(둘 다 선점 기록 전에
  옛 시각을 읽으면) 드물게 이중 조회가 날 수 있다. 겹쳐봐야 우체국 조회 한 번 더일 뿐이라
  과설계하지 않았다. 완전 차단이 필요해지면 시트 클레임 락 도입(참고: TODO의 「시트 클레임 락」).
- **PC 간 시계 오차**에 락 판정이 민감하다(초 단위). 실사용엔 무해한 수준.
- 수동 새로고침은 `tracking_last_auto_refresh`를 **갱신하지 않는다**(사용자 의도 조회는
  드물어 자동과 근접 실행돼도 무방). 필요 시 수동에도 기록하도록 확장 여지.

## 관련 리팩터

- 설정 upsert 로직을 `_write_config_values(ws, updates)`로 분리했다. 락 기록과 기존
  `run_tracking_config_write_worker`가 공용으로 쓴다(중복 제거).
- 타임스탬프 파서 `_parse_ts(s)` 추가('%Y-%m-%d %H:%M:%S', 실패 시 None).
