# 2BS 프리미엄 강의

친구들과 같이 쓰는 강의 수강신청 사이트입니다.

## 사이트 주소

https://eloyprqq.github.io/2bs_story/

## Firebase 설정 (필수)

1. [Firebase Console](https://console.firebase.google.com/) 에서 새 프로젝트 만들기
2. **Authentication → Sign-in method → Google** 켜기
3. **Authentication → Settings → Authorized domains** 에 `eloyprqq.github.io` 추가
4. **Firestore Database** 만들기 (테스트 모드로 시작 가능)
5. **프로젝트 설정 → 일반 → 내 앱 → 웹 `</>`** 추가 후 `firebaseConfig` 복사
6. `index.html` 의 `FIREBASE_CONFIG` 에 붙여넣기
7. GitHub에 push

## 선생님 시크릿

수강신청 이름에 `송율호와김도윤1234` 입력 → 신청자·질문 목록
