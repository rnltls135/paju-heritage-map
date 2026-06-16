GitHub Pages 업로드용 v8.22-5 사진 수정본

이전 GitHub Pages fixed ZIP은 PWA 필수 파일은 들어갔지만,
사진 URL 매칭 데이터가 빠진 잘못된 정리본이었습니다.

이번 v8.22-5는 다음을 모두 포함합니다.
- v8.22_4 사진 URL 매칭 데이터
- Google Drive 사진 509건
- 매칭 유적 134건
- manifest.json
- service-worker.js
- .nojekyll
- cad_topo_tiles 전체

업로드 방법:
1. 기존 GitHub 저장소 폴더 안 파일을 이 ZIP 압축해제본 내용물로 덮어씁니다.
2. GitHub Desktop에서 Summary에 update photo fixed 입력
3. Commit to main
4. Push origin
5. GitHub Pages 주소를 열고 Ctrl+F5 또는 ?v=8225를 붙여 새로고침합니다.

주의:
- 사진 원본은 앱에 포함하지 않았고, Google Drive URL만 들어갑니다.
- 사진이 액박이면 Google Drive 파일 권한을 '링크가 있는 모든 사용자 - 보기 가능'으로 확인해야 합니다.
