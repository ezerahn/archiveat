# -*- coding: utf-8 -*-
"""class-guide.html 전면 개편 — 오늘 확정한 규칙 반영"""
import re, datetime

KST=datetime.timezone(datetime.timedelta(hours=9))
UPDATED=datetime.datetime.now(KST).strftime("%Y.%m.%d %H:%M")
CHAT="https://claude.ai/chat/909e7f28-5718-4bde-8997-e37348632306"

CSS = """
:root{--bg:#FAF5EC;--ink:#241F1B;--sub:#6B6256;--line:#E7DCC8;--card:#FFFDF9;
--coral:#FF5019;--point:#B4A032;--wait:#9A8F7C;--mute:#ADA294;--ok:#5E7360}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:Pretendard,-apple-system,"Apple SD Gothic Neo",sans-serif;
line-height:1.62;-webkit-font-smoothing:antialiased;padding-bottom:70px;font-size:15px}
.wrap{max-width:1000px;margin:0 auto;padding:0 20px}
.topbar{display:flex;justify-content:space-between;align-items:flex-start;padding:18px 0;gap:12px;flex-wrap:wrap}
.pill{display:inline-flex;align-items:center;gap:4px;font-size:13px;text-decoration:none;color:var(--ink);
background:var(--card);border:1px solid var(--line);border-radius:20px;padding:5px 13px;margin-right:8px}
.pill:hover{border-color:var(--point);color:var(--point)}
.crumb{font-size:12px;color:var(--mute);text-align:right;line-height:1.7}
.crumb .meta{font-family:"DM Mono",monospace;font-size:11px}
.head{padding:14px 0 22px;border-bottom:1px solid var(--line);margin-bottom:6px}
h1{font-family:Hahmlet,"Gowun Batang","Noto Serif KR",serif;font-weight:800;font-size:40px;letter-spacing:-.02em}
.lead{font-size:15px;color:var(--sub);margin-top:12px}
.tabs{display:flex;gap:5px;flex-wrap:wrap;margin:18px 0 0;position:sticky;top:0;z-index:5;
background:var(--bg);padding:10px 0 9px;border-bottom:1px solid var(--line)}
.tab{font-family:inherit;font-size:13.5px;font-weight:600;padding:8px 15px;border-radius:10px;
border:1px solid var(--line);background:var(--card);color:var(--wait);cursor:pointer;transition:.12s}
.tab:hover{color:var(--sub)}
.tab.on{background:var(--coral);border-color:var(--coral);color:#fff}
.pane{display:none}
.pane.on{display:block}
.toc{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:11px 15px;margin:16px 0 20px}
.toc a{display:inline-block;font-size:12.5px;color:var(--sub);text-decoration:none;
padding:3px 10px;margin:3px 4px 3px 0;border:1px solid var(--line);border-radius:20px;background:#fff}
.toc a:hover{border-color:var(--coral);color:var(--coral)}
h2{font-family:Hahmlet,"Gowun Batang",serif;font-size:24px;margin:36px 0 6px;padding-top:8px}
h2 .n{font-family:"DM Mono",monospace;font-size:13px;color:var(--coral);margin-right:9px}
.h2d{font-size:13.5px;color:var(--wait);margin-bottom:14px}
h3{font-family:Hahmlet,serif;font-size:17px;margin:22px 0 8px}
.box{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px 19px;margin-bottom:12px}
.box.key{border-left:3px solid var(--coral);border-radius:0 14px 14px 0}
.box.warn{background:#FFF8EC;border-color:#F0DFB8}
p{margin-bottom:9px}
ul{margin:0 0 9px 18px}
li{margin-bottom:5px}
b{font-weight:700}
code{font-family:"DM Mono",monospace;font-size:12.5px;background:#F1EFE9;padding:1px 6px;border-radius:5px}
table{width:100%;border-collapse:collapse;font-size:13.5px;margin:10px 0 12px}
th{background:#F1EFE9;font-size:11.5px;font-weight:700;color:var(--sub);padding:7px 9px;text-align:left;
border:1px solid var(--line);white-space:nowrap}
td{padding:7px 9px;border:1px solid var(--line);vertical-align:top}
td.n{font-family:"DM Mono",monospace;text-align:right;white-space:nowrap}
td.c{text-align:center;white-space:nowrap}
tr.hi td{background:#FFF1EB}
.bd{display:inline-block;font-size:11px;padding:2px 8px;border-radius:20px;white-space:nowrap}
.b-cor{background:#FFE7DF;color:var(--coral);font-weight:500}
.b-mut{background:#F1EFE9;color:#6B6256}
.b-yel{background:#F5B21F;color:#4A3405;font-weight:500}
.b-grn{background:#E4EAE2;color:var(--ok)}
.rule{background:#FFF1EB;border-radius:9px;padding:10px 14px;margin:8px 0;font-size:13.5px}
.rule b{color:var(--coral)}
.foot{margin-top:34px;padding-top:18px;border-top:1px solid var(--line);
font-family:"DM Mono",monospace;font-size:11px;color:var(--wait);text-align:center}
@media(max-width:820px){h1{font-size:29px}.crumb{text-align:left}}
"""

S=[]
def sec(no,title,desc,body,tab="기본"):
    S.append((no,title,desc,body,tab))

sec("01","이 문서를 먼저 읽으세요",
 "새 대화창에서 클래스 작업을 시작할 때 가장 먼저 볼 문서입니다",
"""
<div class="box key">
<p><b>클래스 영역의 모든 규칙이 여기 있습니다.</b> 화면을 만들거나 고칠 때, 새 제휴사 자료를 붙일 때
이 문서의 규칙을 따릅니다. 대화에서 정한 것은 창이 바뀌면 사라지므로, 확정된 규칙은 반드시 여기에 기록합니다.</p>
</div>
<h3>파일 구성</h3>
<table>
<tr><th>구분</th><th>파일</th><th>역할</th></tr>
<tr><td>마스터</td><td><code>class-master-all-v3.xlsx</code></td><td>클래스 회차 (온라인·자체)</td></tr>
<tr><td>마스터</td><td><code>갤러리아_강좌이력.xlsx</code></td><td>제휴사별 아카이브</td></tr>
<tr><td>생성기</td><td><code>gen_class.py</code></td><td>class.html · class-settle-report.html</td></tr>
<tr><td>생성기</td><td><code>gen_list.py</code></td><td><b>class-list.html</b> — 여러 마스터 통합</td></tr>
<tr><td>생성기</td><td><code>gen_partner.py</code></td><td>class-partner.html</td></tr>
<tr><td>생성기</td><td><code>gen_menu.py</code></td><td>class-menu.html — <b>menu-images/ 필수</b></td></tr>
</table>
<div class="rule"><b>실행 순서</b> — gen_class → gen_partner → <b>gen_list(마지막)</b>.
gen_list가 모든 마스터를 읽으므로 반드시 마지막에 돌립니다.</div>
<div class="box warn">
<p><b>gen_menu.py 주의</b> — <code>menu-images/</code> 폴더 없이 돌리면 class-menu.html의 사진 44장이 사라집니다.
안전장치를 넣어 사진이 없으면 중단되며, <code>--force</code>를 붙여야만 강행됩니다.
<b>메뉴 사전을 고칠 일이 없으면 아예 돌리지 마세요.</b></p>
</div>
""","시작")

sec("02","회차ID · 파일명 규칙",
 "2026-07-28 확정. 한번 정하면 바꾸지 않습니다",
"""
<div class="rule"><b>회차ID</b> = <code>[제휴사코드]-[연도2자리][월2자리]-[일련2자리]</code></div>
<table>
<tr><th>제휴사</th><th>코드</th><th>예</th></tr>
<tr><td>경기도지식</td><td class="c">GS</td><td><code>GS-2607-01</code></td></tr>
<tr><td>갤러리아광교</td><td class="c">GL</td><td><code>GL-2511-03</code></td></tr>
<tr><td>클래스콕</td><td class="c">CK</td><td><code>CK-2508-01</code></td></tr>
</table>
<ul>
<li>일련번호는 <b>같은 제휴사·같은 달</b> 안에서 수업일 순으로 01부터</li>
<li><b>ID는 한번 붙이면 바꾸지 않습니다.</b> 제휴사코드는 「붙일 당시의 소속」 표시일 뿐이며,
나중에 제휴사가 바뀌면 ID는 그대로 두고 <code>제휴사</code> 열만 고칩니다</li>
<li>필터·집계는 ID가 아니라 <code>제휴사</code> 열로 합니다. ID 앞자리는 사람이 읽고 부르기 위한 것</li>
</ul>
<div class="rule"><b>파일명</b> = <code>class-[용도]-[회차ID].확장자</code> · 하이픈 제거 · 모든 HTML은 <code>class</code>로 시작</div>
<table>
<tr><th>용도</th><th>예</th></tr>
<tr><td>패킹</td><td><code>class-pack-GS260701.html</code></td></tr>
<tr><td>레시피</td><td><code>class-recipe-GS260701.html</code></td></tr>
<tr><td>커버 이미지</td><td><code>class-cover-GS260701.png</code></td></tr>
</table>
<p>GitHub Pages는 저장소 하나에 모든 영역 파일이 섞이므로 클래스 파일은 전부 <code>class</code>로 시작해야 합니다.
엑셀 마스터는 GitHub에 올리지 않으므로 한글 이름을 유지합니다.</p>
""","규칙")

sec("03","유형 · 테마",
 "유형은 정산 구조로 정하고, 성격은 테마로 분리합니다",
"""
<div class="rule"><b>유형은 이름이 아니라 정산 구조가 정합니다.</b></div>
<table>
<tr><th>명세 행수</th><th>유형</th><th>뜻</th></tr>
<tr><td class="c">3행</td><td><span class="bd b-cor">정규</span></td><td>강사료 + 재료비 2종으로 분리 청구</td></tr>
<tr><td class="c">1행</td><td><span class="bd b-mut">원데이</span></td><td>단일 항목</td></tr>
</table>
<ul>
<li><b>키즈는 유형이 아니라 테마입니다.</b> 아동 대상 수업이라도 1행이면 원데이, 3행이면 정규</li>
<li>테마는 클래스명 앞머리(<code>[키즈]</code> <code>[아동]</code>)에서 자동 추출하며, 한식·베이킹 등은 나중에 마스터에 직접 채웁니다</li>
<li>테마 필터로 「어떤 주제가 잘 열렸나」를 봅니다</li>
</ul>
""","규칙")

sec("04","돈의 흐름 — 매출 · 지출 · 수익",
 "모든 화면이 이 3단 구분을 따릅니다",
"""
<table>
<tr><th>구분</th><th>단계</th><th>뜻</th></tr>
<tr><td rowspan="3" class="c"><b>매출</b></td><td>클래스매출</td><td>수강생이 낸 돈 · 판매가 × 인원</td></tr>
<tr><td>− 제휴사 수수료</td><td>제휴사가 떼는 몫</td></tr>
<tr><td><b>= 제휴사 입금</b></td><td>우리에게 실제로 들어온 돈</td></tr>
<tr><td rowspan="3" class="c"><b>지출</b></td><td>− 선생님 정산</td><td><b>실지급액</b> — 통장에서 실제로 나간 돈</td></tr>
<tr><td>− 재료·배송비</td><td>온라인만 · 우리 부담</td></tr>
<tr><td>− 제휴사 지급</td><td>앞으로 생길 수 있음 (클래스101 등)</td></tr>
<tr class="hi"><td class="c"><b>수익</b></td><td><b>= PK 수익</b></td><td>남는 돈</td></tr>
</table>
<div class="box key">
<p><b>이름 규칙</b> — <code>제휴사 입금</code>은 받는 돈, <code>제휴사 지급</code>은 주는 돈입니다.
<code>제휴사 정산</code>처럼 방향이 안 보이는 말은 쓰지 않습니다.</p>
</div>
<h3>우리 몫 공식 (오프라인 제휴)</h3>
<div class="rule"><b>PK 수익 = 제휴사 입금 − 선생님 정산(실지급) − 재료·배송비</b></div>
<p>모두 <b>실제로 오간 금액</b>입니다. 선생님 정산은 계약 기준이 아니라 <b>통장에서 나간 실지급액</b>을 씁니다
(사업자는 부가세를 얹어서, 프리랜서는 원천세를 떼고 나갑니다).</p>
<h3>참고 — 계약 기준으로 본 배분 구조</h3>
<div class="rule">계약 기준 몫(공급가 매출 대비 %) = <b>25% − 제휴사 수수료율</b></div>
<ul>
<li><b>25%</b> = 공급가 매출을 100%로 봤을 때 선생님 몫 75%를 뺀 나머지</li>
<li>선생님 몫은 <b>수수료율과 무관하게 고정</b>입니다. 제휴 수수료는 전액 우리 몫에서만 빠집니다</li>
<li>갤 0% → 25% · 갤 10% → 15% · <b>갤 30% → −5%(역마진)</b></li>
<li>수수료율을 정할 때 쓰는 <b>참고값</b>입니다. 화면의 PK 수익은 실지급 기준이라 이보다 큽니다</li>
</ul>
""","돈")

sec("05","부가세 — 어디에 쓰고 어디에 안 쓰는가",
 "가장 자주 헷갈리는 지점입니다",
"""
<table>
<tr><th>단계</th><th>VAT 포함</th><th>VAT 제외</th></tr>
<tr><td>클래스매출</td><td class="c">○</td><td class="c">○</td></tr>
<tr><td>제휴사 수수료</td><td class="c">○</td><td class="c">○</td></tr>
<tr><td>제휴사 입금</td><td class="c">○</td><td class="c">○</td></tr>
<tr><td>선생님 정산</td><td class="c">—</td><td class="c">○</td></tr>
<tr><td>재료·배송비</td><td class="c">—</td><td class="c">○</td></tr>
<tr class="hi"><td>PK 수익</td><td class="c">—</td><td class="c">○</td></tr>
</table>
<div class="rule"><b>매출 구간</b>(클래스매출~제휴사 입금)만 VAT 포함·제외가 나뉩니다.
<b>선생님 정산·재료비</b>는 VAT 구분 없이 한 가지 금액입니다.</div>
<ul>
<li>제휴사 수수료는 <b>부가세 포함 매출</b>에 수수료율을 곱합니다 (세금계산서도 그 금액 기준)</li>
<li>선생님 정산은 <b>공급가 매출</b>(부가세 제외)의 75%입니다. 제휴사가 뗀 뒤 금액이 아닙니다</li>
<li>재료비는 <b>실제 구매가(부가세 포함)</b> 그대로 씁니다. 매입세액공제 등 세무 처리는 이 화면에서 다루지 않습니다</li>
<li>손익은 전부 <b>부가세를 뺀 기준</b>으로 봅니다</li>
</ul>
""","돈")

sec("06","선생님 정산 — 계약 기준과 실지급",
 "같은 수업인데 통장 숫자가 다른 이유",
"""
<table>
<tr><th>구분</th><th>계약 기준</th><th>실지급</th><th>차이</th></tr>
<tr><td><span class="bd b-cor">사업자</span></td><td class="n">153,409</td><td class="n">168,750</td><td>부가세 +10% 가산</td></tr>
<tr><td><span class="bd b-mut">프리랜서</span></td><td class="n">239,318</td><td class="n">231,438</td><td>원천세 −3.3% 공제</td></tr>
</table>
<ul>
<li><b>계약 기준</b> = 드리기로 한 금액. 공급가 매출 × 75%. <b>수수료율을 정할 때 참고</b>합니다</li>
<li><b>실지급</b> = 통장에서 실제로 나간 금액. <b>★ 화면의 PK 수익은 이 금액을 씁니다</b></li>
<li>사업자는 부가세를 얹어 보내고 나중에 매입세액공제로 환급받습니다</li>
<li>프리랜서는 원천세를 떼고 보내되, 그 세금도 우리가 국세청에 대납하므로 지갑에서 나가는 총액은 같습니다</li>
</ul>
<h3>원천세 계산</h3>
<div class="rule"><b>단순 3.3%가 아닙니다.</b> 소득세(3%, 10원 절사) + 지방소득세(소득세의 10%, 10원 절사)</div>
<ul>
<li>계산 단위는 <b>지급 묶음 합계</b>이지 회차별이 아닙니다</li>
<li>명세는 회차별로 절사해 넣어놓은 경우가 있어 통장과 10~360원 차이가 납니다. <b>오류가 아닙니다</b></li>
</ul>
<h3>이체 명의</h3>
<div class="box warn">
<p>이체 명의가 선생님 성함과 다를 수 있습니다 — 박소진 → <b>이도현</b>, 일호실 → <b>이민영</b>.
지급내역의 <code>이체 명의</code> 칸에 반드시 기록합니다. 없으면 「이체 내역이 없다」고 헤매게 됩니다.</p>
</div>
""","돈")

sec("07","전체 흐름 — 클래스 하나가 지나는 길",
 "기획부터 정산 완료까지. 화면이 어떻게 갈라지는지",
"""
<div class="rule"><b>수업 상태</b>와 <b>정산 상태</b>는 별개입니다. 섞어 쓰면 「정산완료인데 수업은 언제 했지?」가 됩니다.</div>
<h3>수업 상태 — 「상태」 열</h3>
<table>
<tr><th>단계</th><th>뜻</th><th>클래스 홈에서</th></tr>
<tr><td><span class="bd b-mut">기획중</span></td><td>개설·준비 중</td><td>기획중 그룹</td></tr>
<tr><td><span class="bd b-mut">제안중</span> <span class="bd b-mut">검토중</span></td><td>견적 제출 · 확정 전</td><td>제안중 그룹</td></tr>
<tr><td><span class="bd b-cor">확정</span></td><td>개강 결정</td><td><b>곧 수업</b> 그룹</td></tr>
<tr><td><span class="bd b-cor">모집중</span></td><td>신청 받는 중</td><td>모집 중 그룹</td></tr>
<tr class="hi"><td><span class="bd b-grn">수업완료</span></td><td>수업이 끝남</td><td><b>정산 예정</b> 그룹</td></tr>
<tr><td><span class="bd b-off">폐강</span></td><td>최소 인원 미달 등으로 취소</td><td>안 보임</td></tr>
</table>
<h3>정산 상태 — 「입금」·「지급」 배지</h3>
<table>
<tr><th>구분</th><th>값</th><th>뜻</th></tr>
<tr><td>입금</td><td><span class="bd b-yel">입금대기</span> → <span class="bd b-grn">입금완료</span></td><td>제휴사가 우리에게 정산금을 보냈는가</td></tr>
<tr><td>지급</td><td><span class="bd b-yel">지급대기</span> → <span class="bd b-grn">지급완료</span></td><td>우리가 선생님께 보냈는가</td></tr>
</table>
<h3>화면별로 어디에 나타나나</h3>
<table>
<tr><th>단계</th><th>클래스 홈</th><th>정산 리포트</th><th>클래스 목록</th></tr>
<tr><td>기획중 · 제안중 · 모집중 · 확정</td><td><b>진행 중인 클래스</b></td><td>—</td><td>보임 (개최 필터에 따라)</td></tr>
<tr class="hi"><td>수업완료 + 정산 미완료</td><td><b>정산 예정</b></td><td><b>진행 중 · 정산 대기</b></td><td>보임 · 실적에 집계</td></tr>
<tr><td>수업완료 + 입금·지급 완료</td><td>안 보임</td><td>전체 정산 내역</td><td>보임 · 실적에 집계</td></tr>
<tr><td>폐강</td><td>안 보임</td><td>안 보임</td><td>개최 필터 「폐강」에서</td></tr>
</table>
<div class="box warn">
<p><b>상태는 자동으로 바뀌지 않습니다.</b> 수업일이 지나도 마스터의 「수업상태」를 직접 <code>수업완료</code>로 고쳐야 합니다.
정산 리포트가 <b>「수업일 지남」</b> 배지로 알려주니 그때 고치면 됩니다.</p>
</div>
<h3>지급 예정일</h3>
<table>
<tr><th>구분</th><th>규칙</th></tr>
<tr><td>자체 클래스</td><td>수업일 1~15일 → 당월 25일 / 16~말일 → 익월 10일</td></tr>
<tr class="hi"><td><b>제휴사</b></td><td><b>수업 익월 25일</b> — 익월 10일 정산금 입금을 확인한 뒤 지급</td></tr>
</table>
<p>갤러리아 자금 흐름 3단계 — ① 월 초(5~9일) 우리 사이트 결제분을 갤러리아에 송금 →
② <b>월 10일</b> 갤러리아가 수수료 뗀 정산금 입금 → ③ <b>익월 25일</b> 선생님께 지급.</p>
""","상태")

sec("08","실적 집계 기준",
 "무엇을 실적으로 셀 것인가",
"""
<div class="rule">실적은 <b>수업완료</b> 회차만 집계합니다.</div>
<h3>★ 상태는 두 갈래 — 섞지 말 것</h3>
<table>
<tr><th>구분</th><th>값</th><th>어디에</th></tr>
<tr><td><b>수업 상태</b></td><td>기획중 → 제안중·검토중 → 확정 → 모집중 → <b>수업완료</b> / 폐강</td><td>「상태」 열</td></tr>
<tr><td><b>정산 상태</b></td><td>입금대기 ↔ 입금완료 · 지급대기 ↔ 지급완료</td><td>「입금」·「지급」 배지</td></tr>
</table>
<p><b>「정산완료」는 수업 상태가 아닙니다.</b> 수업이 끝났으면 <code>수업완료</code>이고,
정산이 끝났는지는 입금·지급 배지로 봅니다. 제휴사 마스터도 <code>입금상태</code>·<code>지급상태</code> 칸(W·X열)을 그대로 읽습니다.</p>
<ul>
<li><code>확정</code> <code>모집중</code> 등 아직 안 끝난 회차는 <b>예정</b>으로 분리하고 대시보드에 「예정 N회 제외」로 표시합니다</li>
<li>상태를 <code>수업완료</code>로 바꾸면 <b>자동으로 실적에 들어옵니다</b></li>
<li><code>폐강</code>은 개최 필터에서 걸러지며, 기본값은 <b>개최</b>입니다</li>
</ul>
<h3>인원 세는 법</h3>
<ul>
<li>정규는 한 회차가 3항목으로 쪼개지므로 <b>강사료 행의 인원</b>이 실제 수강생입니다</li>
<li>재료비 두 항목의 인원 합 = 강사료 인원 (41회차 전부 일치)</li>
</ul>
""","상태")

sec("09","용어 사전 — 이 말만 씁니다",
 "같은 것을 두 이름으로 부르면 나중에 반드시 헷갈립니다",
"""
<div class="rule">왼쪽 표현은 <b>쓰지 않습니다.</b> 오른쪽으로만 씁니다.</div>
<h3>돈</h3>
<table>
<tr><th>쓰지 않는 말</th><th>쓰는 말</th><th>뜻</th></tr>
<tr><td>총매출 · 클래스 매출액</td><td><b>클래스매출</b></td><td>수강생이 낸 돈 · 판매가 × 인원</td></tr>
<tr><td>갤 수수료 · 발주처 수수료</td><td><b>제휴사 수수료</b></td><td>제휴사가 떼는 몫</td></tr>
<tr><td>갤러리아 정산 · 갤정산 · 우리 수령 · 실제 입금 · 정산 수령액</td><td><b>제휴사 입금</b></td><td>제휴사가 <u>우리에게</u> 보낸 돈</td></tr>
<tr><td>제휴사 정산</td><td><b>제휴사 입금</b> / <b>제휴사 지급</b></td><td>방향이 보이게 나눠 씁니다</td></tr>
<tr><td>선생님 지급 · 강사료 · 쌤 정산</td><td><b>선생님 정산</b></td><td>선생님께 드리는 돈</td></tr>
<tr class="hi"><td><b>순이익 · 회사 이익 · 우리 몫 · 마진</b></td><td><b>PK 수익</b></td><td>다 빼고 남는 돈</td></tr>
<tr><td>원가 · 재료비(단독)</td><td><b>재료·배송비</b></td><td>패킹비 + 배송비 + 재료비</td></tr>
</table>
<h3>부가세</h3>
<table>
<tr><th>쓰지 않는 말</th><th>쓰는 말</th></tr>
<tr><td>부가세 포함 · VAT포함(붙여쓰기)</td><td><b>VAT 포함</b></td></tr>
<tr><td>부가세 제외 · 공급가 · 공급가액(표 안에서)</td><td><b>VAT 제외</b></td></tr>
</table>
<p><code>공급가액</code>은 <b>세금계산서 발행 기준</b>을 적을 때만 씁니다. 표 열 이름에는 <code>VAT 제외</code>로 통일합니다.</p>
<h3>합계</h3>
<table>
<tr><th>쓰지 않는 말</th><th>쓰는 말</th><th>어디에</th></tr>
<tr><td>합계 · 총계 · 소계</td><td><b>계</b></td><td>열 이름 (예: 선생님 정산 · 계)</td></tr>
<tr><td>—</td><td><b>합계</b></td><td>표 맨 아래 합계 <u>행</u>에만</td></tr>
<tr><td>—</td><td><b>회차 합계</b></td><td>펼침 표의 합계 행</td></tr>
</table>
<h3>분류</h3>
<table>
<tr><th>쓰지 않는 말</th><th>쓰는 말</th><th>비고</th></tr>
<tr><td>단일 · 단기 · 특강</td><td><b>원데이</b></td><td>명세 1행짜리</td></tr>
<tr><td>키즈(유형으로)</td><td>유형은 <b>정규 / 원데이</b></td><td>키즈는 <b>테마</b> 열로</td></tr>
<tr><td>발주처</td><td><b>제휴사</b></td><td>클래스 마스터 열 이름은 「발주처」지만 화면에는 제휴사</td></tr>
<tr><td>정산완료(수업 상태로)</td><td>수업 상태는 <b>수업완료</b></td><td>정산 여부는 입금·지급 배지로</td></tr>
</table>
<h3>사람</h3>
<table>
<tr><th>쓰지 않는 말</th><th>쓰는 말</th></tr>
<tr><td>일호실 (닉네임)</td><td><b>이민영</b></td></tr>
<tr><td>강사 · 쌤</td><td><b>선생님</b></td></tr>
</table>
<div class="box warn">
<p><b>축약하지 않습니다.</b> <code>갤정산</code>·<code>쌤정산</code>처럼 줄이면 나중에 본 사람이 못 알아봅니다.
길어도 <code>제휴사 입금</code>·<code>선생님 정산</code>으로 씁니다.</p>
</div>
""","화면")

sec("10","화면 표기 규칙",
 "만들 때마다 다시 묻지 않도록",
"""
<h3>표 머리글 3단</h3>
<table>
<tr><th>1단</th><td>클래스</td><td>매출</td><td>지출</td><td>수익</td></tr>
<tr><th>2단</th><td></td><td>클래스매출 · 제휴사 입금</td><td>선생님 정산 · 재료·배송비</td><td>PK 수익</td></tr>
<tr><th>3단</th><td>회차ID~인원</td><td>VAT 포함 · VAT 제외 · 수수료</td><td>정산기준 · 계</td><td>VAT 제외</td></tr>
</table>
<h3>지켜야 할 것</h3>
<ul>
<li><b>굵은 글씨는 계산에 영향을 주는 금액</b>에만 씁니다. 어떤 값의 근거가 다른 열에 있으면 그 열을 같은 굵기로 맞춥니다</li>
<li><b>이름은 앞 장 「용어 사전」을 따릅니다</b> — 같은 것을 두 이름으로 부르지 않습니다</li>
<li>그룹 경계에는 <b>굵은 세로선</b>, 열 사이에는 얇은 선. 1단 그룹 안에는 선을 넣지 않습니다</li>
<li>항목이 여러 개인 행에는 <b>▾ 아이콘</b>을 붙여 눌러볼 수 있음을 표시합니다</li>
<li>펼친 블록은 <b>왼쪽 코랄 세로선</b>으로 부모 행에 붙이고, 합계 줄을 넣어 접힌 값과 이어짐을 보입니다</li>
<li>금액 아래 작은 회색 글씨로 <b>계산 근거</b>를 답니다 (비율·기준)</li>
</ul>
<h3>배지 색</h3>
<table>
<tr><th>색</th><th>뜻</th><th>예</th></tr>
<tr><td><span class="bd b-cor">코랄</span></td><td>진행·주요</td><td>정규 · 사업자 · 개최</td></tr>
<tr><td><span class="bd b-mut">무채색</span></td><td>일반·완료 전</td><td>원데이 · 프리랜서</td></tr>
<tr><td><span class="bd b-grn">회녹</span></td><td>완료</td><td>수업완료 · 정산완료</td></tr>
<tr><td><span class="bd b-yel">노랑</span></td><td>확인 필요</td><td>미확인 · 대상 미상</td></tr>
</table>
""","화면")

sec("11","새 제휴사 자료를 붙일 때",
 "클래스콕 · 경기도지식 등",
"""
<div class="rule">공통 12항목만 맞추면 <code>gen_list.py</code>에 파일명 한 줄 추가로 붙습니다.</div>
<table>
<tr><th>#</th><th>항목</th><th>비고</th></tr>
<tr><td class="c">1</td><td>회차ID</td><td>제휴사코드-연월-일련</td></tr>
<tr><td class="c">2</td><td>제휴사</td><td>필터·집계의 기준</td></tr>
<tr><td class="c">3</td><td>수업일</td><td></td></tr>
<tr><td class="c">4</td><td>클래스명</td><td></td></tr>
<tr><td class="c">5</td><td>구분</td><td>온라인 / 오프라인</td></tr>
<tr><td class="c">6</td><td>유형</td><td>명세 행수로 판정</td></tr>
<tr><td class="c">7</td><td>선생님</td><td></td></tr>
<tr><td class="c">8</td><td>개최</td><td>개최 / 폐강</td></tr>
<tr><td class="c">9</td><td>인원</td><td>정규는 강사료 행 기준</td></tr>
<tr><td class="c">10</td><td>클래스매출</td><td>VAT 포함</td></tr>
<tr><td class="c">11</td><td>제휴사 수수료</td><td>없으면 0</td></tr>
<tr class="hi"><td class="c">12</td><td><b>PK 수익</b></td><td>VAT 제외 · 비교 기준</td></tr>
</table>
<p>부가 항목 — 제휴사 입금 · 선생님 정산(계약/실지급) · 재료·배송비 · 지급방식 · 테마 · 수업상태</p>
""","확장")

sec("12","미해결 · 확인 대기",
 "다음에 이어서 할 것",
"""
<table>
<tr><th>항목</th><th>내용</th></tr>
<tr><td>갤러리아 이체 6건</td><td>대응 회차 없음 · 지급내역 U-01~U-06 · 1,808,641원</td></tr>
<tr><td>경기도지식 지급방식</td><td>사업자/프리랜서 미확인 · 마스터에 칸 없음</td></tr>
<tr><td>패킹비·배송비</td><td>인당 5,000원이 실비인지 기준값인지 미확인</td></tr>
<tr><td>클래스콕</td><td>자료 확보 후 제휴사 탭 추가</td></tr>
<tr><td>정산 리포트 2단계</td><td>세금계산서 발행·원천세 신고 포함</td></tr>
</table>
""","확장")

TABS=[("시작","시작하기"),("규칙","기본 규칙"),("돈","돈의 흐름"),
      ("상태","전체 흐름"),("화면","화면 표기"),("확장","확장·미해결")]
tabbtn="".join(f'<button class="tab{" on" if i==0 else ""}" data-t="{k}">{lab}</button>'
               for i,(k,lab) in enumerate(TABS))
panes=""
for i,(k,lab) in enumerate(TABS):
    items=[x for x in S if x[4]==k]
    toc="".join(f'<a href="#s{no}">{no} {t}</a>' for no,t,_,_,_ in items) if len(items)>1 else ""
    body="".join(
      f'<h2 id="s{no}"><span class="n">{no}</span>{t}</h2><p class="h2d">{d}</p>{b}'
      for no,t,d,b,_ in items)
    panes+=f'<div class="pane{" on" if i==0 else ""}" data-t="{k}">'+ \
           (f'<div class="toc">{toc}</div>' if toc else '')+body+'</div>'

JS = """
document.querySelector('.tabs').addEventListener('click',function(ev){
  var b=ev.target.closest('.tab'); if(!b)return;
  var t=b.dataset.t;
  this.querySelectorAll('.tab').forEach(function(x){x.classList.toggle('on',x===b);});
  document.querySelectorAll('.pane').forEach(function(p){p.classList.toggle('on',p.dataset.t===t);});
  window.scrollTo({top:0,behavior:'smooth'});
});
document.querySelectorAll('.toc a').forEach(function(a){
  a.addEventListener('click',function(ev){
    ev.preventDefault();
    var el=document.querySelector(this.getAttribute('href'));
    if(el)window.scrollTo({top:el.offsetTop-70,behavior:'smooth'});
  });
});
"""

doc=f"""<!DOCTYPE html><html lang="ko"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<meta name="version" content="v4"><meta name="updated" content="{UPDATED}">
<title>클래스 설계 가이드</title>
<link rel="apple-touch-icon" href="app-icon.png"><link rel="icon" href="app-icon.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Hahmlet:wght@400;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css" rel="stylesheet">
<style>{CSS}</style></head><body><div class="wrap">

<div class="topbar">
<div><a class="pill" href="class.html">← 클래스 홈</a><a class="pill" href="{CHAT}" target="_blank" rel="noopener">💬 작업 대화</a></div>
<div class="crumb">AI 워크스페이스 › 클래스 › 설계 가이드<br><span class="meta">업데이트 {UPDATED} · v3</span></div>
</div>

<div class="head">
<h1>클래스 설계 가이드</h1>
<p class="lead">클래스 영역의 <b>확정된 규칙</b>을 담은 문서입니다. 새 대화창에서 작업을 시작할 때 이 문서를 먼저 읽고,
새로 정해진 규칙은 반드시 여기에 기록합니다.</p>
</div>

<div class="tabs">{tabbtn}</div>
{panes}

<div class="foot">클래스 설계 가이드 · {UPDATED} · v4</div>
</div>
<script>{JS}</script></body></html>"""

open("class-guide.html","w",encoding="utf-8").write(doc)
import os
print(f"class-guide.html v4 — {len(S)}개 장 · {os.path.getsize('class-guide.html')/1024:.0f}KB")
