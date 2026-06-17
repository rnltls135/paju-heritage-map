GitHub Pages 업로드용 v8.22-6 PWA 설치 안정화본

변경 내용:
- 사진 URL 매칭 데이터 유지: 509건 / 134개 유적
- CAD 수치지형도 타일 포함
- manifest.json의 PWA id를 안정값(/paju-heritage-map/)으로 고정
- service-worker 캐시명 갱신
- 설치 버튼 안내 문구 개선
- .nojekyll 유지

업로드:
1. 이 ZIP을 압축 해제
2. GitHub 저장소 폴더 안 기존 파일을 모두 덮어쓰기
3. GitHub Desktop에서 Commit
4. Push origin
5. https://rnltls135.github.io/paju-heritage-map/?v=8226 으로 접속

설치가 계속 안 되면:
- 기존 설치된 '파주 유적지도' 앱 삭제
- Chrome에서 사이트 데이터/캐시 삭제
- 다시 접속 후 10~20초 기다린 뒤 Chrome 메뉴 ⋮ > 앱 설치 또는 홈 화면에 추가
