# typing_solution

수행방법:
1. make_folder.exe를 실행하면 C드라이브 안에 typing_solution 폴더가 생성됨.(C:\typing_solution)

2. typing_solution 폴더 안에 검사하고자 하는 hwp 파일 '한 개'를 넣음.

3. hangul_ver2_exe.exe 를 실행

4. 'hwp 수식 문법 검사기'에서 start를 누르고 팝업창이 뜨면 '접근 허용'이나 '모두 허용' 클릭

4-1. 한글 파일이 뜨고 탐색하는 동안 마우스 클릭 금지

5. 탐색이 끝나면 'hwp 수식 문법 검사기'에서 next 버튼을 눌러 틀린 문법을 검사

6. 틀린 문법을 모두 찾았으면 '끝'이라고 출력되고 C:\typing_solution\log\(년월일_시분초)_log.txt 파일에 탐색한 기록이 저장된다.

검사 목록
"""
x->INF -> x `->` INF
x rarrow INF -> x `rarrow` INF
rm 없을때 A빼고 전부 뒤에 `붙이기 it붙은거 포함
log 밑, 지수 앞에 `붙이기
!=1 -> != `1
=1 -> =`1
- -> -`
+ -> +`
< -> <`
> -> >`
% -> %'
TIMES -> TIMES `
/ -> /`
``vert `` 나중에 팔요하면 구현
(-1,3) or (-1,`3) -> (-1,``3)
"""
