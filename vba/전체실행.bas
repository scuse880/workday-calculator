Attribute VB_Name = "전체실행"
Option Explicit

Public Sub main_전체실행()

    ' 1. 근무상황목록 반영
    main_근무상황목록.main_근무상황목록

    ' 2. 출장목록(개인별) 반영
    main_출장목록개인별.main_출장목록_개인별

    ' 3. 방학중근무일지 반영
    main_방학중근무일지.main_방학중근무일지

    ' 4. 정액분 판정
    main_정액분계산.main_정액분계산

    MsgBox "전체 처리가 완료되었습니다.", vbInformation

End Sub
