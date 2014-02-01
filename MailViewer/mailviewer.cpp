///////////////////////////////////////////////////////////////////////////////////////
// Mail Viewer Version 0.07
// 作成者:砂倉瑞希(Angelic Software)
// 更新バージョン:0.07(2007年12月27日)
// http://www.天使のお茶会.com/
//////////////////////////////////////////////////////////////////////////////////////

#pragma warning(disable : 4996)

#include "mailviewer.h"

// WinMain関数
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInst, LPSTR lpszCmd, int nCmdShow){

	MSG msg;
	HWND hFind;
	HACCEL hAc;
	INITCOMMONCONTROLSEX ic;	// コモンコントロール

	// 二重起動禁止
	if((hFind = FindWindow(NULL, "Mail Viewer")) != NULL) {
		SetForegroundWindow(hFind);
		return(0);
	}

	// メール格納ディレクトリの初期化
	if(!InitFolder()){
		MessageBox(NULL, "メールフォルダの作成に失敗しました。", "エラー", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// nMail.dllの初期化
	NMailSetParameter(0, TEMP_MAX, 0);
	NMailInitializeWinSock();
	
	// コモンコントロールの初期化
	ic.dwSize = sizeof(INITCOMMONCONTROLSEX);
	ic.dwICC = ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&ic);

	// アプリケーションの初期化
	if(!InitApp(hInstance)){
		return FALSE;
	}

	// グローバル変数にコピー
	hInst = hInstance;

	// ウィンドウの作成
	if(!InitInstance(hInstance, nCmdShow)){
		return FALSE;
	}

	// アクセラレータキーとHTMLヘルプの初期化
	hAc = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
	HtmlHelp(NULL, NULL, HH_INITIALIZE, (DWORD)&dwCookie);

	while(bRet = GetMessage(&msg, NULL, 0, 0) != 0){
		// GetMessageがエラーのとき
		if(bRet == -1){
			break;
		}
		else{
			if(TranslateAccelerator(hParent, hAc, &msg) == 0){
				if(!HtmlHelp(NULL, NULL, HH_PRETRANSLATEMESSAGE, (DWORD)&msg)){
					TranslateMessage(&msg);
					DispatchMessage(&msg);
				}
			}
		}
	}

	return(msg.wParam);
}

// アプリケーションの初期化
ATOM InitApp(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	// ウィンドウクラスの設定
	wcex.hInstance = hInstance;
	wcex.lpszClassName = szWinName;
	wcex.lpfnWndProc = WndProc;
	wcex.style = CS_HREDRAW | CS_VREDRAW;
	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
	wcex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
	wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
	wcex.lpszMenuName = MAKEINTRESOURCE(IDR_MENU1);
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	
	return(RegisterClassEx(&wcex));
}

// ウィンドウの作成
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;

	// ウィンドウの作成
	hWnd = CreateWindow(szWinName, "Mail Viewer", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);

	if(!hWnd){
		return FALSE;
	}

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
	
	// グローバル変数にコピー
	hParent = hWnd;

	return TRUE;
}

// ウィンドウプロシージャ
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int Ret;
	static int iStatusWy;
	static HWND hEdit, hList;	// コントロールハンドル
	static int sortsubno[NO_OF_SUBITEM] = {UP};
	static int isortsubno;

	RECT rc;					// RECT構造体
	POINT pt;					// POINT構造体
	DWORD dwStyle;				// ウィンドウスタイル
	LPNMHDR lpnmhdr;			// ヘッダハンドラ
	LPNMLISTVIEW lpnmlv;		// リストビューハンドラ
	HMENU hMenu, hPMenu;		// メニューハンドル
	SORTDATA SortData;			// リストのソート用構造体
	NMLISTVIEW *pNMLV;

	switch(msg){
		case WM_CREATE:
            hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "EDIT", "", WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_WANTRETURN | WS_VSCROLL | WS_HSCROLL | ES_AUTOVSCROLL | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, (HMENU)ID_EDIT, hInst, NULL);
			SendMessage(hEdit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), 0);
			hList = CreateWindowEx(WS_EX_CLIENTEDGE, WC_LISTVIEW, "", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT, 0, 0, 0, 0, hWnd, (HMENU)ID_LISTVIEW, hInst, NULL);
			dwStyle = ListView_GetExtendedListViewStyle(hList);
			dwStyle |= LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP;
			ListView_SetExtendedListViewStyle(hList, dwStyle);
			InsertColumn(hList);
			hStatus = CreateStatusWindow(WS_CHILD | WS_VISIBLE, "", hWnd, ID_STATUS);
			GetWindowRect(hStatus, &rc);
			iStatusWy = rc.bottom - rc.top;

			// レジストリの読み込み
			LoadRegKey();

			// 起動時のバグを回避
			No = -1;

			// ローカルに保存しているメール一覧を表示
			SetLocalMail(hList);

			SortData.hwndList = hList;
			isortsubno = 2;
			SortData.isortSubItem = isortsubno;
			SortData.iUPDOWN = sortsubno[isortsubno];

			if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
				MessageBox(hWnd, "ソートに失敗しました。", "エラー", MB_OK | MB_ICONSTOP);
			}
			break;
		case WM_SIZE:
			MoveWindow(hList, 0, 0, LOWORD(lParam), (HIWORD(lParam) - iStatusWy) / 2, TRUE);
			MoveWindow(hEdit, 0, (HIWORD(lParam) - iStatusWy) / 2, LOWORD(lParam), (HIWORD(lParam) - iStatusWy) / 2, TRUE);
			sb_size = LOWORD(lParam) - 120;
			SendMessage(hStatus, WM_SIZE, wParam, lParam);
			break;
		case WM_NOTIFY:
			lpnmhdr = (LPNMHDR)lParam;

			// 新しいコード
			if(lpnmhdr->hwndFrom == hList){
				switch(lpnmhdr->code){
					// シングルクリックの時
					case NM_CLICK:
						lpnmlv = (LPNMLISTVIEW)lParam;
						No = lpnmlv->iItem;
						
						// クリックした番号が-1の時はメールを表示しない
						if(No == -1){
							break;
						}
						
						// メールデータファイル名を取得する
						ListView_GetItemText(hList, No, 4, szSelMailFileName, sizeof(szSelMailFileName));

						// メールファイルを開く
						Read_Mail(szSelMailFileName, hEdit);
					break;
					case NM_RCLICK:
						lpnmlv = (LPNMLISTVIEW)lParam;
						No = lpnmlv->iItem;
						
						// クリックした番号が-1の時はメニューを表示しない
						if(No == -1){
							break;
						}

						// 右クリックメニューを開く
						GetCursorPos(&pt);
						CreatePopupMenu();
						hPMenu = LoadMenu(hInst, MAKEINTRESOURCE(IDR_MENU2));
						hMenu = GetSubMenu(hPMenu, 0);
						TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
						DestroyMenu(hMenu);
					break;
					case LVN_COLUMNCLICK:
						pNMLV = (NMLISTVIEW *)lParam;
						isortsubno = pNMLV->iSubItem;

						if(sortsubno[isortsubno] == UP){
							sortsubno[isortsubno] = DOWN;
						}
						else{
							sortsubno[isortsubno] = UP;
						}

						SortData.hwndList = hList;
						SortData.isortSubItem = isortsubno;
						SortData.iUPDOWN = sortsubno[isortsubno];
						
						if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
							MessageBox(hWnd, "ソートに失敗しました。", "エラー", MB_OK | MB_ICONSTOP);
						}

					break;
				}
			}
            break;
		// メニューを選択した時の説明表示(ステータスバー)
		case WM_MENUSELECT:
			switch (LOWORD(wParam)){
				case IDM_OPEN:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[0]);
					break;
				case IDM_SAVE:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[1]);
					break;
				case IDM_PAGEINF:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[2]);
					break;
				case IDM_PRINT:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[3]);
					break;
				case IDM_EXIT:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[4]);
					break;
				case IDM_NEW:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[5]);
					break;
				case IDM_RCVMAIL:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[6]);
					break;
				case IDM_REPLY:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[7]);
					break;
				case IDM_DELETE:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[8]);
					break;
				case IDM_HEADER:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[9]);
					break;
				case IDM_SET:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[10]);
					break;
				case IDM_HELP:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[11]);
					break;
				case IDM_ABOUT:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szTipText[12]);
					break;
				default:
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
					break;
			}
			break;
		// メニューを抜けた時の処理
		case WM_EXITMENULOOP:
			SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
			break;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDM_NEW:
					// 新規
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG4), hWnd, (DLGPROC)NewMailProc);
					break;
				case IDM_OPEN:
					// 開く
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG5), hWnd, (DLGPROC)AttachFileProc);
					break;
				case IDM_SAVE:
					// 保存
					if(No >= 0){
						WriteMailFile(hWnd);
					}
					break;
				case IDM_PRINT:
					// 印刷
					if(No >= 0){
						// メールの件名を取得
						ListView_GetItemText(hList, No, 0, szSelMailSubject, sizeof(szSelMailSubject));
						wsprintf(strMsg, "「%s」を印刷しますか？", szSelMailSubject);

						Ret = MessageBox(hWnd, strMsg, "印刷の確認", MB_YESNO | MB_ICONQUESTION);
						if(Ret == IDYES){
							MailPrint(hWnd);
						}
					}
					break;
				case IDM_RCVMAIL:
					// メールの受信
					// エディットボックスの内容を削除する
					SetWindowText(hEdit, "");

					// メール受信部分のサブルーチン
					RcvMail(hWnd, hList);

					SortData.hwndList = hList;
					SortData.isortSubItem = isortsubno;
					SortData.iUPDOWN = sortsubno[isortsubno];

					if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
						MessageBox(hWnd, "ソートに失敗しました。", "エラー", MB_OK | MB_ICONSTOP);
					}
					break;
				case IDM_REPLY:
					// メールの返信
					if(No >= 0){
						bReply = TRUE;
						DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG4), hWnd, (DLGPROC)NewMailProc);
					}
					break;
				case IDM_DELETE:
					// メールの削除
					if(No >= 0){
						// テキストボックスを初期化
						SetWindowText(hEdit, "");

						// メールデータファイル名を取得する
						ListView_GetItemText(hList, No, 4, szSelMailFileName, sizeof(szSelMailFileName));
						ListView_GetItemText(hList, No, 0, szSelMailSubject, sizeof(szSelMailSubject));

						wsprintf(strMsg, "「%s」を削除しますか？", szSelMailSubject);
						Ret = MessageBox(hWnd, strMsg, "削除の確認", MB_YESNO | MB_ICONQUESTION);

						if(Ret == IDYES){
							// 選択したメールを削除
							if(!DeleteFile(szSelMailFileName)){
								// 削除に失敗したとき
								MessageBox(hWnd, "削除に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
								break;
							}
							else{
								// 削除が出来たとき
								MessageBox(hWnd, "メールを削除しました", "Mail Viewer", MB_OK | MB_ICONINFORMATION);
							}

							// 受信メール一覧の更新
							SetLocalMail(hList);

							SortData.hwndList = hList;
							SortData.isortSubItem = isortsubno;
							SortData.iUPDOWN = sortsubno[isortsubno];
							
							if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
								MessageBox(hWnd, "ソートに失敗しました。", "エラー", MB_OK | MB_ICONSTOP);
							}
						}
					}
					break;
				case IDM_HEADER:
					// ヘッダ表示
					if(No >= 0){
						DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG3), hWnd, (DLGPROC)HeaderProc);
						SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
					}
					break;
				case IDM_SET:
					// メールの設定
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG1), hWnd, (DLGPROC)MailSetProc);
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
					break;
				case IDM_EXIT:
					SendMessage(hWnd, WM_CLOSE, 0, 0);
					break;
				case IDM_HELP:
					HtmlHelp(hWnd, "mailviewer.chm::/00.html > mailviewer01", HH_DISPLAY_TOC, NULL);
					break;
				case IDM_ABOUT:
					VersionNo = NMailGetVersion();
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG2), hWnd, (DLGPROC)AboutProc);
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
					break;
			}
			return 0;
		case WM_CLOSE:
			// HTMLヘルプが開いていたら閉じる
			if(HtmlHelp(hWnd, "mailviewer.chm", HH_GET_WIN_HANDLE, (DWORD)"mailviewer01")){
				HtmlHelp(NULL, NULL, HH_CLOSE_ALL, 0);
			}

			// レジストリの保存
			SaveRegKey();

			// ウィンドウハンドルを破棄
			DestroyWindow(hWnd);
			DestroyWindow(hStatus);
			DestroyWindow(hEdit);
			break;
		case WM_DESTROY:
			HtmlHelp(NULL, NULL, HH_UNINITIALIZE, dwCookie);
			NMailEndWinSock();
			PostQuitMessage(0);
			break;
		default:
			return(DefWindowProc(hWnd, msg, wParam, lParam));
	}
	return (0L);
}

// メールの設定ダイアログ
LRESULT CALLBACK MailSetProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hUser, hPass, hName, hMailAddr, hSmtp, hPop3, hAPopCheck, hPSmtpCheck, hSMailDelCheck;
	
	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hUser = GetDlgItem(hDlg, IDC_EDIT1);
			hPass = GetDlgItem(hDlg, IDC_EDIT2);
			hName = GetDlgItem(hDlg, IDC_EDIT3);
			hMailAddr = GetDlgItem(hDlg, IDC_EDIT4);
			hSmtp = GetDlgItem(hDlg, IDC_EDIT5);
			hPop3 = GetDlgItem(hDlg, IDC_EDIT6);
			hAPopCheck = GetDlgItem(hDlg, IDC_CHECK1);
			hPSmtpCheck = GetDlgItem(hDlg, IDC_CHECK2);
			hSMailDelCheck = GetDlgItem(hDlg, IDC_CHECK3);
			Edit_SetText(hUser, szUserName);
			Edit_SetText(hPass, szPass);
			Edit_SetText(hName, szName);
			Edit_SetText(hMailAddr, szMailAddr);
			Edit_SetText(hSmtp, szSmtpServer);
			Edit_SetText(hPop3, szPopServer);
			SendMessage(hAPopCheck, BM_SETCHECK, (WPARAM)bApop, 0L);
			SendMessage(hPSmtpCheck, BM_SETCHECK, (WPARAM)bPSmtp, 0L);
			SendMessage(hSMailDelCheck, BM_SETCHECK, (WPARAM)bSMailDel, 0L);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDOK:
					Edit_GetText(hUser, szUserName, sizeof(szUserName));
					Edit_GetText(hPass, szPass, sizeof(szPass));
					Edit_GetText(hName, szName, sizeof(szName));
					Edit_GetText(hMailAddr, szMailAddr, sizeof(szMailAddr));
					Edit_GetText(hSmtp, szSmtpServer, sizeof(szSmtpServer));
					Edit_GetText(hPop3, szPopServer, sizeof(szPopServer));

					// APOPのチェック
					if(IsDlgButtonChecked(hDlg, IDC_CHECK1) == BST_CHECKED){
						bApop = TRUE;
					}
					else{
						bApop = FALSE;
					}

					// POP Before SMTPのチェック
					if(IsDlgButtonChecked(hDlg, IDC_CHECK2) == BST_CHECKED){
						bPSmtp = TRUE;
					}
					else{
						bPSmtp = FALSE;
					}

					// サーバにメールを残すのチェック
					if(IsDlgButtonChecked(hDlg, IDC_CHECK3) == BST_CHECKED){
						bSMailDel = TRUE;
					}
					else{
						bSMailDel = FALSE;
					}

					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDCANCEL);
					return TRUE;
				default:
					return FALSE;
			}
			default:
				return FALSE;
	}
	return TRUE;
}

// バージョン情報ダイアログ
LRESULT CALLBACK AboutProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hStatic;
	char strVerNo[256];
	double dNo;

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hStatic = GetDlgItem(hDlg, IDC_STATIC1);
			dNo = itod(VersionNo, 0.01);		// 整数を小数点付きの数値に変換
			sprintf(strVerNo, "%0.2f", dNo);	// wsprintfだと無理なのでsprintfに変更
			SetWindowText(hStatic, strVerNo);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDOK:
					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDOK);
					return TRUE;
				default:
					return FALSE;
			}
			default:
				return FALSE;
	}
	return TRUE;
}

// ヘッダ情報表示ダイアログ
LRESULT CALLBACK HeaderProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hEdit = GetDlgItem(hDlg, IDC_EDIT1);
			SetWindowText(hEdit, szHeaderMsg);
			wsprintf(strMsg, "\"%s\"のヘッダ情報", szHeader);
			SetWindowText(hDlg, strMsg);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDOK:
					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDOK);
					return TRUE;
				default:
					return FALSE;
			}
			default:
				return FALSE;
	}
	return TRUE;
}

// メール作成ダイアログ
LRESULT CALLBACK NewMailProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hAddr, hAttach, hSubject, hEdit, hAtchBtn;
	char szFileTitle[SEND_MAX], szRe[TEMP_MAX], szResInfo[TEMP_MAX];
	int recnt, NoTitle;
	BOOL bResult;

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hAtchBtn = GetDlgItem(hDlg, IDC_BUTTON1);
			hAddr = GetDlgItem(hDlg, IDC_EDIT1);
			hSubject = GetDlgItem(hDlg, IDC_EDIT2);
			hAttach = GetDlgItem(hDlg, IDC_EDIT3);
			hEdit = GetDlgItem(hDlg, IDC_EDIT4);
			if(bReply == TRUE){
				SetWindowText(hAddr, RepFrom);
				if(recnt = strstr(RepSubject, "Re:") == 0){
					sprintf(szRe, "Re: %s", RepSubject);
					SetWindowText(hSubject, szRe);
				}
				else{
					SetWindowText(hSubject, RepSubject);
				}

				sprintf(szResInfo, "\r\n\r\n----Original Message----\r\nFrom : %s\r\nDate : %s\r\nSubject : %s\r\n\r\n", RepFrom, RepDate, RepSubject);
				strcat(szResInfo, strText);

				if(szResInfo == NULL){
					MessageBox(hDlg, "エラーが発生しました。", "エラー", MB_OK | MB_ICONSTOP);
				}
				SetWindowText(hEdit, szResInfo);
				bReply = FALSE;
			}
			if(strcmp(szFileData, "") != 0){
				strcpy(szFileData, "");
			}
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDOK:
					// 内容チェック(宛先、件名空欄)機能
					Edit_GetText(hAddr, szToData, sizeof(szToData));
					
					if(strcmp(szToData, "") == 0 || strstr(szToData, "@") == 0 || strstr(szToData, ".") == 0){
						MessageBox(hDlg, "宛先が入力されていませんか、不正なメールアドレスです", "エラー", MB_OK | MB_ICONSTOP);
						return FALSE;
					}

					Edit_GetText(hSubject, szSubjectData, sizeof(szSubjectData));

					// ※未承諾広告送信不可
					if(strcmp(szSubjectData, "※未承諾広告") == 0 || strcmp(szSubjectData, "広告") == 0 || strcmp(szSubjectData, "承諾") == 0){
						MessageBox(hDlg, "このソフトは広告メールを送信できなくしています", "エラー", MB_OK | MB_ICONSTOP);
						return FALSE;
					}

					if(strcmp(szSubjectData, "") == 0){
						NoTitle = MessageBox(hDlg, "件名が入力されていませんが送信してよろしいですか？\nはいの場合は件名は(無題)で送信されます", "件名未入力", MB_YESNO | MB_ICONINFORMATION);
						if(NoTitle == IDNO){
							return FALSE;
						}
						else{
							strcpy(szSubjectData, "(無題)");
						}
					}
					
					Edit_GetText(hEdit, szBodyData, sizeof(szBodyData));
					Edit_GetText(hAttach, szFileData, sizeof(szFileData));

					// Pop Before SMTPが選択されている時はPOP3に接続してから送信する
					if(bPSmtp == TRUE){
						// POP3サーバに接続して切断する
						bResult = SmtpPop(s, szPopServer, szUserName, szPass);

						// 戻り値がFALSEの時は接続失敗
						if(bResult == FALSE){
							// POP3サーバ接続失敗
							MessageBox(hDlg, "POP3サーバに接続できませんでした。\nPOP3サーバの設定を再確認してください。", "エラー", MB_OK | MB_ICONSTOP);
							return FALSE;
						}
					}

					// メールを送信する
					Send_Mail(szSmtpServer, szToData, szSubjectData, szBodyData, szFileData);

					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDCANCEL);
					return TRUE;
				case IDM_CLOSE:
					EndDialog(hDlg, IDCANCEL);
					return TRUE;
				case IDM_ATTATCH:
					SendMessage(hDlg, WM_COMMAND, IDC_BUTTON1, 0);
					return TRUE;
				case IDM_SENDMAIL:
					SendMessage(hDlg, WM_COMMAND, IDOK, 0);
					return TRUE;
				case IDM_UNDO:
					SendMessage(hEdit, WM_UNDO, 0, 0);
					return TRUE;
				case IDM_CUT:
					SendMessage(hEdit, WM_CUT, 0, 0);
					return TRUE;
				case IDM_COPY:
					SendMessage(hEdit, WM_COPY, 0, 0);
					return TRUE;
				case IDM_PASTE:
					SendMessage(hEdit, WM_PASTE, 0, 0);
				case IDM_DEL:
					SendMessage(hEdit, WM_CLEAR, 0, 0);
					break;
					return TRUE;
				case IDM_ALLSEL:
					SendMessage(hEdit, EM_SETSEL, (WPARAM)0, (LPARAM)-1);
					return TRUE;
				case IDC_BUTTON1:
					strcpy(szFileTitle, "");
					SelectAttachFile(hDlg, szFileTitle, szFileData);
					SetWindowText(hAttach, szFileData);
					break;
				case IDM_HELP:
					HtmlHelp(hDlg, "mailviewer.chm::/00.html > mailviewer01", HH_DISPLAY_TOC, NULL);
					break;
				default:
					return TRUE;
			}
			default:
				return FALSE;
	}

	return TRUE;
}

// 添付ファイルを展開して保存ダイアログ
LRESULT CALLBACK AttachFileProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	HWND hList, hFDEdit, hFBtn, hFDBtn;
	
	BROWSEINFO bi;
	ITEMIDLIST *lpid;
	HRESULT hr;
	LPMALLOC pMalloc = NULL;				// IMallocへのポインタ
	BOOL bDir = FALSE;

	HANDLE hFile;							// ファイルハンドル
	HGLOBAL hMem;							// グローバルハンドル
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// バイト数
	char *lpszBuf;							// ファイルバッファ
	char f_id[TEMP_MAX], *PertId;			// 分割メールのID
	int no;									// 添付ファイルの個数
	int ret;								// 戻り値

	// 添付メール展開用変数
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX], body[TEMP_MAX];
	char FileName[TEMP_MAX], Temp[NMAIL_ATTACHMENT_TEMP_SIZE];

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hList = GetDlgItem(hDlg, IDC_LIST1);
			hFDBtn = GetDlgItem(hDlg, IDC_FILEBTN);
			hFBtn = GetDlgItem(hDlg, IDC_BUTTON1);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
			case IDC_BUTTON1:
				AttachFileOpen(hDlg, szAtFile, szAtFileName);

				// ファイル名が入力されていないとき
				if(strlen(szAtFileName) == 0){
					return FALSE;
				}

				// ファイルを開く
				hFile = CreateFile(szAtFile, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
				
				if(hFile == INVALID_HANDLE_VALUE){
					MessageBox(hDlg, "ファイルをオープンできません", "エラー", MB_OK | MB_ICONSTOP);
					return FALSE;
				}
				
				// ファイルサイズを調べる
				dwFSize = GetFileSize(hFile, &dwFSizeHigh);
				
				if(dwFSizeHigh != 0){
					MessageBox(hDlg, "ファイルが大きすぎます", "エラー", MB_OK | MB_ICONSTOP);
					CloseHandle(hFile);
					return FALSE;
				}

				// ファイルを読み込むためのメモリ領域を確保
				hMem = GlobalAlloc(GHND, dwFSize + 1);
				
				if(hMem == NULL){
					MessageBox(hDlg, "メモリを確保できません", "エラー", MB_OK | MB_ICONSTOP);
					CloseHandle(hFile);
					return FALSE;
				}
				
				// メモリ領域をロックしてファイルを読み込む
				lpszBuf = (char *)GlobalLock(hMem);
				ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
				lpszBuf[dwFSize] = '\0';
				
				// szAtFileDataのメモリ領域を確保する
				szAtFileData = (char *)malloc(sizeof(char) * dwFSize);

				if(szAtFileData == NULL){
					MessageBox(hDlg, "メモリ確保に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
					return FALSE;
				}

				// メールファイルの中身をコピーする
				strcpy((char *)szAtFileData, lpszBuf);
				
				// ファイルを閉じる
				CloseHandle(hFile);
				GlobalUnlock(hMem);
				GlobalFree(hMem);
				
				// ファイルが添付されているかを確認する
				if((no = NMailAttachmentFileStatus(szAtFileData, f_id, TEMP_MAX)) != NMAIL_NO_ATTACHMENT_FILE){
					if(no == 1){
						PertId = f_id;
					}
					hList = GetDlgItem(hDlg, IDC_LIST1);
					if(no == 0 || PertId == f_id){
						SendMessage(hList, LB_INSERTSTRING, (WPARAM)0, (LPARAM)szAtFile);
						SendMessage(hList, LB_SETCURSEL, (WPARAM)0, 0L);
						EnableWindow(GetDlgItem(hDlg, IDOK), TRUE);
					}
				}
				else{
					// 添付ファイルがないのでメッセージを出して終了
					MessageBox(hDlg, "このメールに、添付ファイルはありません", "添付ファイルなし", MB_OK |MB_ICONINFORMATION);
				}

				return TRUE;

			case IDC_FILEBTN:
				memset(&bi, 0, sizeof(BROWSEINFO));
				bi.hwndOwner = hDlg;
				bi.lpfn = SHMyProc;
				bi.ulFlags = BIF_EDITBOX | BIF_STATUSTEXT | BIF_VALIDATE;
				bi.lpszTitle = "展開先ディレクトリ指定";
				lpid = SHBrowseForFolder(&bi);
				
				if(lpid == NULL){
					return FALSE;
				}
				else{
					hr = SHGetMalloc(&pMalloc);
					if(hr == E_FAIL){
						MessageBox(hDlg, "SHGetMalloc Error", "エラー", MB_OK | MB_ICONSTOP);
						return FALSE;
					}
					SHGetPathFromIDList(lpid, szDir);
					if(szDir[strlen(szDir) - 1] != '\\'){
						strcat(szDir, "\\");
					}

					hFDEdit = GetDlgItem(hDlg, IDC_EDIT1);

					// 展開先フォルダをセットする
					SetWindowText(hFDEdit, szDir);

					pMalloc->Free(lpid);
					pMalloc->Release();
					bDir = TRUE;
				}
				return TRUE;

				case IDOK:
					// 添付ファイルの展開
					ret = NMailAttachmentFileFirst(Temp, subject, date, from, header, body, szDir, FileName, szAtFileData, NULL);
					if(ret == NMAIL_SUCCESS){
						NMailAttachmentFileClose(Temp);
						wsprintf(strMsg, "%sに%sを展開しました。", szDir, FileName);
						MessageBox(hDlg, strMsg, "Mail Viewer", MB_OK | MB_ICONINFORMATION);
					}
					// メモリの解放
					free(szAtFileData);
					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDOK);
					return TRUE;
				default:
					return FALSE;
			}
			default:
				return FALSE;
	}
	return TRUE;
}

// リストに件名、差出人、日付のカラムを挿入する
void InsertColumn(HWND hList)
{
    LVCOLUMN lvcol;

    lvcol.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvcol.fmt = LVCFMT_LEFT;
    lvcol.cx = 300;
    lvcol.pszText = "件名";
    lvcol.iSubItem = 0;
    ListView_InsertColumn(hList, 0, &lvcol);

    lvcol.cx = 250;
    lvcol.pszText = "差出人";
    lvcol.iSubItem = 1;
    ListView_InsertColumn(hList, 1, &lvcol);

    lvcol.cx = 155;
    lvcol.pszText = "日付";
    lvcol.iSubItem = 2;
    ListView_InsertColumn(hList, 2, &lvcol);

    lvcol.cx = 50;
    lvcol.pszText = "サイズ";
    lvcol.iSubItem = 3;
    ListView_InsertColumn(hList, 3, &lvcol);

	// メールのファイルパスを格納(非表示)
    lvcol.cx = 0;
    lvcol.pszText = "";
    lvcol.iSubItem = 4;
    ListView_InsertColumn(hList, 4, &lvcol);

    return;
}

// リストに件名、差出人、日付を挿入する
void InsertItem(HWND hList, int iItemNo, int iSubNo, char *lpszText)
{
	LVITEM item;
	int lvitem;

	memset(&item, 0, sizeof(LVITEM));
	lvitem = iItemNo - 1;	// lvitemの初期値にiItemNo - 1の値を代入

	if(iSubNo == 0){
		item.mask = LVIF_TEXT | LVIF_PARAM;
	}
	else{
		item.mask = LVIF_TEXT;
	}

	item.pszText = lpszText;
	item.iItem = iItemNo;
	item.iSubItem = iSubNo;

	if(iSubNo == 0){
		item.lParam = iItemNo;
		lvitem = ListView_InsertItem(hList, &item);
	}
	else{
		item.iItem = lvitem;
		ListView_SetItem(hList, &item);
	}

	return;
}

// メールのリストを表示する
int List_Mail(SOCKET s, char *host, char *id, char *pass, HWND hList, HWND hStatus)
{
	int count, max, rcount, newcnt, msize = 0;
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX];
	char strMsg[1024], szSize[1024];
	char szlYear[4], szlMonth[2], szlDay[2], szlHour[2], szlMinute[2], szlSecond[2];
	char szName[256];
	LPSTR lpUidl, lpFilePath;
	BOOL bFile;

	max = 0;
	rcount = 0;
	newcnt = 0;

	// 認証
	if((max = NMailPop3Authenticate(s, id, pass, bApop)) >= 0){
		// サーバにあるメールの数が0以上のとき
		if(max > 0 ){
			// プログレスパー表示
			hProgress = CreateWindowEx(0, PROGRESS_CLASS, "", WS_CHILD | WS_VISIBLE, sb_size, 5, 100, 16, hStatus, (HMENU)ID_PROGRESS, hInst, NULL);
			SendMessage(hProgress, PBM_SETRANGE, 0, MAKELONG(0, max));
		}
		else{
			// プログレスバーを削除
			DestroyWindow(hProgress);
		}

		// 受信しているメールの数分ヘッダを読み出す
		for(count = 0; count < max; count++){

			// プログレスバーを増加
			SendMessage(hProgress, PBM_SETPOS, count + 1, 0);

			if(NMailPop3GetMailStatus(s, count + 1, subject, date, from, header, FALSE) >= 0){

				// うまくいかなかったのでサイズ取得方法を変更
				if((msize = NMailPop3GetMailSize(s, count + 1)) >= 0){
					sprintf(szSize, "%d", msize);
				}
				else{
					wsprintf(szSize, "0");
				}

				// 取得した日付を分解
				mid(date, szlYear, 1, 4);
				mid(date, szlMonth, 6, 2);
				mid(date, szlDay, 9, 2);
				mid(date, szlHour, 12, 2);
				mid(date, szlMinute, 15, 2);
				mid(date, szlSecond, 18, 2);

				// UIDLのサイズを確保
				lpUidl = (char *)malloc(sizeof(char) * 1024);
				lpFilePath = (char *)malloc(sizeof(char) * MAX_PATH);

				if(lpUidl == NULL){
					free(lpUidl);
					break;
				}
				
				if(lpFilePath == NULL){
					free(lpFilePath);
					break;
				}

				// カレントディレクトリを取得する
				GetCurrentDirectory(MAX_PATH, lpFilePath);
				strcat(lpFilePath, "\\inbox\\");

				// UIDLを取得する
				NMailPop3GetUidl(s, count + 1, lpUidl, 1024);

				// リスト用のファイル名を作成
				wsprintf(szName, "%s%s%s%s%s%s_%s.dat", szlYear, szlMonth, szlDay, szlHour, szlMinute, szlSecond, lpUidl);

				strcat(lpFilePath, szName);
				bFile = FileExists(lpFilePath);

				// ファイルが存在しない(新規メール)とき
				if(bFile == FALSE){
					// 処理中のメッセージの表示
					wsprintf(strMsg, "%d件目を処理中", count + 1);
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)strMsg);

					// メールデータを保存
					Store_Mail(s, count + 1, lpFilePath);
				}
				else{
					// 受信済みの場合は受信済みカウントを増加する
					rcount++;
				}

				// サーバにメールを残すチェックがはずされているとき
				if(bSMailDel == FALSE){
					// サーバのメールを削除する
					NMailPop3DeleteMail(s, count + 1);
				}

				free(lpUidl);
				free(lpFilePath);

			}
			else{
				break;
			}
			MessagePump();
		}
	}
	else{
		// 認証エラーの場合(maxにはエラー値が格納されている)
		return max;
	}

	if(max > 0){
		// プログレスバーを削除
		DestroyWindow(hProgress);
	}

	// 新規メールから受信済みメールの数を引く
	newcnt = max - rcount;

	// 新着メールがあるとき
	if(newcnt > 0){
		// 保存されたメールの一覧を表示する
		SetLocalMail(hList);
	}

	// ステータスメッセージの初期化
	strcpy(strMsg, "");

	// 最大件数が0件(メールがない)の場合
	if(newcnt == 0){
		// メールがサーバにない(または既読メールしかない)時
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"メールは届いていません");
	}
	else{
		// メールがある場合は件数表示
		wsprintf(strMsg, "%d件のメールが届いています", newcnt);
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)strMsg);
	}

	return max;
}

//////////////////////////////////////////////////////////////////////////
// Send_Mail
// 内容　：メールを送信する
// 戻り値：なし
//////////////////////////////////////////////////////////////////////////
void Send_Mail(char *host, char *to, char *subject, char *body, char *atatch)
{
	char from[SEND_MAX];

	// fromの部分のフォーマット作成
	sprintf(from, "%s <%s>", szName, szMailAddr);

	// 本文の終わりに\r\nを付ける
	strcat(body, "\r\n");

    // メール送信
	if(NMailSmtpSendMail(host, to, NULL, NULL, from, subject, body, "X-MAILER: Mail Viewer Version 0.0.7", atatch, 0) >= 0){
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"メールを送信しました");
	}
	else{
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"メール送信に失敗しました");
	}
}

// ダイアログを中央に表示する
BOOL CenterWindow(HWND hwndChild, HWND hwndParent)
{
	BOOL bResult;
	RECT rcChild, rcParent, rcWorkArea;
	int	 wChild, hChild, wParent, hParent;
	int	 xNew, yNew;

	GetWindowRect(hwndChild,&rcChild);
	wChild = rcChild.right - rcChild.left;
	hChild = rcChild.bottom - rcChild.top;

	if(hwndParent){
		GetWindowRect(hwndParent,&rcParent);
	}
	else{
		GetWindowRect(GetDesktopWindow(),&rcParent);
	}
	
	wParent = rcParent.right - rcParent.left;
	hParent = rcParent.bottom - rcParent.top;

	bResult = SystemParametersInfo(SPI_GETWORKAREA,sizeof(RECT),&rcWorkArea,0);

	if(!bResult){
		rcWorkArea.left = rcWorkArea.top = 0;
		rcWorkArea.right = GetSystemMetrics(SM_CXSCREEN);
		rcWorkArea.bottom = GetSystemMetrics(SM_CYSCREEN);
	}

	xNew = rcParent.left + ((wParent-wChild) / 2);

	if(xNew < rcWorkArea.left){
		xNew = rcWorkArea.left;
	}
	else if((xNew+wChild) > rcWorkArea.right){
		xNew = rcWorkArea.right - wChild;
	}

	yNew = rcParent.top  + ((hParent-hChild) / 2);

	if(yNew < rcWorkArea.top){
		yNew = rcWorkArea.top;
	}
	else if((yNew+hChild) > rcWorkArea.bottom){
		yNew = rcWorkArea.bottom - hChild;
	}

	return SetWindowPos(hwndChild, HWND_TOP, xNew, yNew, 0, 0, SWP_NOSIZE);
}

// 添付ファイルの処理
void SelectAttachFile(HWND hDlg, char *lpstrFile, char *lpstrTitle)
{
	OPENFILENAME ofn;

	memset(&ofn, 0, sizeof(OPENFILENAME));

	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hDlg;
	ofn.lpstrFilter = "すべてのファイル(*.*)\0*.*\0\0";
	ofn.lpstrFile = lpstrFile;
	ofn.lpstrFileTitle = lpstrTitle;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = "txt";
	ofn.lpstrTitle = "添付ファイルの選択";
	ofn.nMaxFileTitle = MAX_PATH;
	GetOpenFileName(&ofn);

	return;
}

// メールファイルを開く
void AttachFileOpen(HWND hDlg, char *lpstrFile, char *lpstrTitle)
{
	OPENFILENAME ofn;	// 開くダイアログ構造体

	memset(&ofn, 0, sizeof(OPENFILENAME));

	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hDlg;
	ofn.lpstrFilter = "Mail Viewer(*.mvr)\0*.mvr\0メールファイル(*.eml)\0*.eml\0テキストファイル(*.txt)\0*.txt\0すべてのファイル(*.*)\0*.*\0";
	ofn.lpstrFile = lpstrFile;
	ofn.lpstrFileTitle = lpstrTitle;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = "mvr";
	ofn.lpstrTitle = "メールファイルを開く";
	ofn.nMaxFileTitle = MAX_PATH;
	GetOpenFileName(&ofn);

	return;
}

// メールをテキストで出力
int WriteMailFile(HWND hWnd)
{
	OPENFILENAME ofn;
	HANDLE hFile;
	DWORD dwAccBytes;

	memset(&ofn, 0, sizeof(OPENFILENAME));
	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hWnd;
	ofn.lpstrFilter = "Mail Viewer(*.mvr)\0*.mvr\0メールファイル(*.eml)\0*.eml\0テキストファイル(*.txt)\0*.txt\0すべてのファイル(*.*)\0*.*\0";
	ofn.lpstrFile = szFileName;
	ofn.lpstrFileTitle = szFile;
	ofn.nFilterIndex = 1;
	ofn.nMaxFile = sizeof(szFileName);
	ofn.nMaxFileTitle = sizeof(szFile);
	ofn.Flags = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = "txt";

	if(GetSaveFileName(&ofn) == 0){
		return -1;
	}

	hFile = CreateFile(szFileName, GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	WriteFile(hFile, szMailData, strlen(szMailData) + 1, &dwAccBytes, NULL);
	
	if(CloseHandle(hFile) == 0){
		MessageBox(hWnd, "ファイルを保存できませんでした", "エラー", MB_OK | MB_ICONSTOP);
	}
	return 0;
}

//////////////////////////////////////////////////////////////////////////
// itod(int to double)
// 内容　：整数を小数点付きの数値に変換
// 引数　：num　 ・・・ 整数(123のように記述/int)
// 　　　　point ・・・ 小数点の位置(0.01などで指定/double)
// 戻り値：小数点がついた値(double)、-1が返る場合はエラー
//////////////////////////////////////////////////////////////////////////
double itod(int num, double point){

	double value;

	if(num > 0){
		value = num * point;
	}
	else{
		value = -1;
	}
	return value;
}

//////////////////////////////////////////////////////////////////////////
// MailPrint(Mail Print)
// 内容　：メールの内容を印刷する(簡易印刷)
//////////////////////////////////////////////////////////////////////////
int MailPrint(HWND hWnd){
	PRINTDLG   pd;
	DOCINFO    di;
	TEXTMETRIC tm;
	RECT       rct;

	char szPrnDat[SEND_MAX];

	memset(&pd, 0, sizeof(PRINTDLG));

	pd.lStructSize = sizeof(PRINTDLG);
	pd.hwndOwner   = hWnd;
	pd.hDevMode    = NULL;
	pd.hDevNames   = NULL;
	pd.Flags       = PD_USEDEVMODECOPIESANDCOLLATE | PD_RETURNDC | PD_NOPAGENUMS | PD_NOSELECTION | PD_HIDEPRINTTOFILE;

	pd.nMinPage    = 1;
	pd.nMaxPage    = 1;
	pd.nToPage     = 1;
	pd.nCopies     = 1;

	if(PrintDlg(&pd) == 0){
		return -1;
	}

	memset(&di, 0, sizeof(DOCINFO));

	di.cbSize       = sizeof(DOCINFO);
	di.lpszDocName = RepSubject;

	StartDoc(pd.hDC, &di);
	StartPage(pd.hDC);
	GetTextMetrics(pd.hDC, &tm);
	GetClientRect(hWnd, &rct);
	sprintf(szPrnDat, "\r\n送信時刻：%s \r\n送信者：%s \r\n件名：%s \r\n ------------------------------------------------------------ \r\n %s", RepDate, RepFrom, RepSubject, strText);
	rct.left   = 10;
	rct.right  = tm.tmHeight * strlen(szPrnDat);
	rct.bottom = 10 + (tm.tmHeight * strlen(szPrnDat));
	DrawText(pd.hDC, szPrnDat, -1, &rct, DT_LEFT);

	EndPage(pd.hDC);
	EndDoc(pd.hDC);
	DeleteDC(pd.hDC);

	return(0);

}

// 添付ファイル格納先フォルダを選択する
int CALLBACK SHMyProc(HWND hWnd, UINT msg, LPARAM lp1, LPARAM lp2)
{
    char szFolder[MAX_PATH];
    char szBuf[1024];
    ITEMIDLIST *lpid;

    switch (msg) {
        case BFFM_SELCHANGED:
            lpid = (ITEMIDLIST *)lp1;
            SHGetPathFromIDList(lpid, szFolder);
            wsprintf(szBuf, "現在「%s」が選択されています", szFolder);
            SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, (LPARAM)szBuf);
            return 0;
        case BFFM_VALIDATEFAILED:
            MessageBox(hWnd, "不正なフォルダです", "エラー", MB_OK | MB_ICONSTOP);
            return 1;

    }
    return 0;
}

// POP Before SMTP用のPOP3認証
BOOL SmtpPop(SOCKET s, char *host, char *id, char *pass)
{
	BOOL bSuccess;

	// POP3サーバに接続
	if((s = NMailPop3Connect(host)) != INVALID_SOCKET){
		// POP3サーバの認証
		if(NMailPop3Authenticate(s, id, pass, bApop) >= 0){
			// 0以上ならTRUEを返す
			bSuccess = TRUE;
		}
		else{
			// -1以下ならFALSEを返す
			bSuccess = FALSE;
		}
		// POP3サーバから抜ける
		NMailPop3Close(s);
	}
	else{
		// INVALID_SOCKETの時はFALSEを返す
		bSuccess = FALSE;
	}
	return bSuccess;
}

// メールデータを保存する
int SaveMailData(LPSTR lpFilePath)
{
	HANDLE hFile;
	DWORD dwAccBytes;
	hFile = CreateFile(lpFilePath, GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	
	// ファイルが作成できなかったとき
	if(hFile == INVALID_HANDLE_VALUE){
		return -1;
	}

	WriteFile(hFile, szMailData, strlen(szMailData) + 1, &dwAccBytes, NULL);
	
	if(CloseHandle(hFile) == 0){
		MessageBox(NULL, "ファイルを保存できませんでした", "エラー", MB_OK | MB_ICONSTOP);
		return -1;
	}
	return 0;
}

// List_Mailの時にメールデータを保存する
void Store_Mail(SOCKET s, int no, LPSTR lpFilePath/*lpUidl*/)
{
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX];
	int  size;

	// メールのサイズ分確保してファイルに書き込む
	if((size = NMailPop3GetMailSize(s, no)) >= 0){
		if((body =(char *)malloc(size)) != NULL){
			if(NMailPop3GetMail(s, no, subject, date, from, header, body, NULL, NULL) >= 0){
				sprintf(szMailData, "%s\r\n%s", header, body);
				SaveMailData(lpFilePath);
			}
			else{
				MessageBox(NULL, "エラーが発生しました", "エラー", MB_OK | MB_ICONSTOP);
			}
			free(body);
		}
	}
}

// MID関数
void mid(char *ukeru, char *kaesu, int a, int b)
{
	int	k, n ;
	n = strlen(ukeru);
	
	if((0 < a && a <= n) && (0 < b && b <= n) && (a + b - 1 <= n)){
		for(k = a; k < a + b;  k++){
			kaesu[k-a] = ukeru[k-1];
		}
		kaesu[b] = '\0' ;
	}else{
		kaesu[0] = '\0' ;
	}
}

// ローカルに保存されているメールを取得する
void SetLocalMail(HWND hListL)
{
	WIN32_FIND_DATA wfd;
	HANDLE hFind;
	HANDLE hFile;							// ファイルハンドル
	HGLOBAL hMem;							// グローバルハンドル
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// バイト数
	char szfname[MAX_PATH + 5];				// ファイル名
	char *lpszBuf;							// ファイルバッファ
	int iListNo;

	// 添付メール展開用変数
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX], body[TEMP_MAX];
	char FileName[TEMP_MAX], Temp[NMAIL_ATTACHMENT_TEMP_SIZE], szSize[TEMP_MAX];

	// ListViewの描画を止める
	SetWindowRedraw(GetDesktopWindow(), FALSE);

	// リストの初期化
	ListView_DeleteAllItems(hListL);
	GetCurrentDirectory(MAX_PATH, szfname);
	strcat(szfname, "\\inbox\\*.dat");
	iListNo = 0;

	// メールファイルを探す
	hFind = FindFirstFile(szfname, &wfd);

	// 検索が失敗した場合(ファイルがない場合)
	if(hFind == INVALID_HANDLE_VALUE){
		FindClose(hFind);
		return;
	}
	
	// ファイルがなくなるまでリストに追加
	do{
		// カレントディレクトリを取得する
		GetCurrentDirectory(MAX_PATH, szfname);
		strcat(szfname, "\\inbox\\");

		// 検索ファイル名をコピー
		strcat(szfname, wfd.cFileName);

		// ファイルを開く
		hFile = CreateFile(szfname, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		
		if(hFile == INVALID_HANDLE_VALUE){
			MessageBox(NULL, "ファイルをオープンできません", "エラー", MB_OK | MB_ICONSTOP);
			return;
		}
		
		// ファイルサイズを調べる
		dwFSize = GetFileSize(hFile, &dwFSizeHigh);
		
		if(dwFSizeHigh != 0){
			MessageBox(NULL, "ファイルが大きすぎます", "エラー", MB_OK | MB_ICONSTOP);
			CloseHandle(hFile);
			return;
		}
		
		// ファイルを読み込むためのメモリ領域を確保
		hMem = GlobalAlloc(GHND, dwFSize + 1);
		
		if(hMem == NULL){
			MessageBox(NULL, "メモリを確保できません", "エラー", MB_OK | MB_ICONSTOP);
			CloseHandle(hFile);
			return;
		}
		
		// メモリ領域をロックしてファイルを読み込む
		lpszBuf = (char *)GlobalLock(hMem);
		ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
		lpszBuf[dwFSize] = '\0';
		
		// szAtFileDataのメモリ領域を確保する
		szAtFileData = (char *)malloc(sizeof(char) * dwFSize);
		
		if(szAtFileData == NULL){
			MessageBox(NULL, "メモリ確保に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
			return;
		}
		
		// メールファイルの中身をコピーする
		strcpy((char *)szAtFileData, lpszBuf);
		
		// ファイルを閉じる
		CloseHandle(hFile);
		GlobalUnlock(hMem);
		GlobalFree(hMem);
		
		// 件名、受信日付、差出人の初期化
		strcpy(subject, "");
		strcpy(date, "");
		strcpy(from, "");

		// 添付ファイルの展開(実際はメール本文とヘッダの分割のみ)
		NMailAttachmentFileFirst(Temp, subject, date, from, header, body, NULL, FileName, szAtFileData, NULL);

		// メールサイズを変換する
		wsprintf(szSize, "%d", dwFSize);

		// 件名が空白の時は(件名なし)を表示する
		if(strlen(subject) > 0){
			InsertItem(hListL, iListNo, 0, subject);
		}
		else{
			InsertItem(hListL, iListNo, 0, "(件名なし)");
		}
		InsertItem(hListL, iListNo + 1, 1, from);
		InsertItem(hListL, iListNo + 1, 2, date);
		InsertItem(hListL, iListNo + 1, 3, szSize);
		InsertItem(hListL, iListNo + 1, 4, szfname);

		free(szAtFileData);

		iListNo++;

	}while(FindNextFile(hFind, &wfd) == TRUE);

	FindClose(hFind);

	// ListViewの描画を再開する
	SetWindowRedraw(GetDesktopWindow(), TRUE);

	return;

}

// リストで選択した件名に対応するメールデータを読み出す
void Read_Mail(char *szFName, HWND hEdit)
{
	HANDLE hFile;							// ファイルハンドル
	HGLOBAL hMem;							// グローバルハンドル
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// バイト数
	char *lpszBuf;							// ファイルバッファ
	
	// 添付メール展開用変数
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX], body[TEMP_MAX];
	char FileName[TEMP_MAX], Temp[NMAIL_ATTACHMENT_TEMP_SIZE];
	
	// 返信用変数の初期化
	strcpy(RepSubject, "");			// 返信用の件名作成用変数の初期化
	strcpy(RepFrom, "");			// 返信用の宛先用変数の初期化
	strcpy(RepDate, "");			// 返信用の日付用変数の初期化
	strcpy(RepSubject, subject);	// 返信用の件名作成用変数
	strcpy(RepFrom, from);			// 返信用の宛先用変数
	strcpy(RepDate, date);			// 返信用の日付用変数

	// ファイルを開く
	hFile = CreateFile(szFName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	
	if(hFile == INVALID_HANDLE_VALUE){
		MessageBox(NULL, "ファイルをオープンできません", "エラー", MB_OK | MB_ICONSTOP);
		return;
	}
	
	// ファイルサイズを調べる
	dwFSize = GetFileSize(hFile, &dwFSizeHigh);
	
	if(dwFSizeHigh != 0){
		MessageBox(NULL, "ファイルが大きすぎます", "エラー", MB_OK | MB_ICONSTOP);
		CloseHandle(hFile);
		return;
	}
	
	// ファイルを読み込むためのメモリ領域を確保
	hMem = GlobalAlloc(GHND, dwFSize + 1);
	
	if(hMem == NULL){
		MessageBox(NULL, "メモリを確保できません", "エラー", MB_OK | MB_ICONSTOP);
		CloseHandle(hFile);
		return;
	}
	
	// メモリ領域をロックしてファイルを読み込む
	lpszBuf = (char *)GlobalLock(hMem);
	ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
	lpszBuf[dwFSize] = '\0';
	
	// szAtFileDataのメモリ領域を確保する
	szAtFileData = (char *)malloc(sizeof(char) * dwFSize);
	
	if(szAtFileData == NULL){
		MessageBox(NULL, "メモリ確保に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
		return;
	}
	
	// メールファイルの中身をコピーする
	strcpy((char *)szAtFileData, lpszBuf);
	
	// ファイルを閉じる
	CloseHandle(hFile);
	GlobalUnlock(hMem);
	GlobalFree(hMem);
	
	// 件名、受信日付、差出人の初期化
	strcpy(subject, "");
	strcpy(date, "");
	strcpy(from, "");

	// 添付ファイルの展開(実際はメール本文とヘッダの分割のみ)
	NMailAttachmentFileFirst(Temp, subject, date, from, header, body, NULL, FileName, szAtFileData, NULL);

	// 読み込んだ本文を表示
	sprintf(strText, "%s", body);
	strcpy(szMailData, szAtFileData);
	SetWindowText(hEdit, strText);
	strcpy(szHeaderMsg, header);

	// 件名が存在するとき
	if(strlen(subject) > 0){
		strcpy(szHeader, subject);
	}
	else{
		strcpy(szHeader, "(件名なし)");
	}

}

// ファイルの存在確認
BOOL FileExists(LPSTR lpFileName)
{
	HANDLE hFile;			// ファイルハンドル
	BOOL bFile;				// ファイルフラグ
	WIN32_FIND_DATA	wfd;	// WIN32_FIND_DATA構造体

	// ファイルフラグを初期化
	bFile = FALSE;

	// ファイルを検索する
	hFile = FindFirstFile(lpFileName, &wfd);

	// ファイルが存在するとき
	if(hFile != INVALID_HANDLE_VALUE){
		// ファイルフラグをTRUEにする
		bFile = TRUE;
	}

	// ファイルの検索を閉じる
	FindClose(hFile);

	return bFile;
}

// レジストリの読み込み
void LoadRegKey()
{
	HKEY hRegKey;							// レジストリハンドラ
	unsigned long datasz, Result, datatype;	// レジストリ用の変数

	// レジストリを開く
	RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\Angelic Software\\MailViewer", 0, "MailViewer", 0, KEY_ALL_ACCESS, NULL, &hRegKey, &Result);

	// 初めてレジストリに登録するとき
	if(Result == REG_CREATED_NEW_KEY){
		// レジストリへの新規作成処理
		RegSetValueEx(hRegKey, "Smtp Address", 0, REG_SZ, (LPBYTE)szSmtpServer, strlen(szSmtpServer)+1);
		RegSetValueEx(hRegKey, "Pop3 Address", 0, REG_SZ, (LPBYTE)szPopServer, strlen(szPopServer)+1);
		RegSetValueEx(hRegKey, "Mail Address", 0, REG_SZ, (LPBYTE)szMailAddr, strlen(szMailAddr)+1);
		RegSetValueEx(hRegKey, "User Name", 0, REG_SZ, (LPBYTE)szUserName, strlen(szUserName)+1);
		RegSetValueEx(hRegKey, "From Name", 0, REG_SZ, (LPBYTE)szName, strlen(szName)+1);
		RegSetValueEx(hRegKey, "Password", 0, REG_SZ, (LPBYTE)szPass, strlen(szPass)+1);
		RegSetValueEx(hRegKey, "APop", 0, REG_DWORD, (LPBYTE)&bApop, sizeof(DWORD));
		RegSetValueEx(hRegKey, "Pop Before SMTP", 0, REG_DWORD, (LPBYTE)&bPSmtp, sizeof(DWORD));
		RegSetValueEx(hRegKey, "Server Mail Del", 0, REG_DWORD, (LPBYTE)&bSMailDel, sizeof(DWORD));
	}
	else{
		// レジストリからの読出し処理
		datasz = 256;
		RegQueryValueEx(hRegKey, "Smtp Address",NULL,&datatype,(LPBYTE)szSmtpServer,&datasz);
		datasz = 256;
		RegQueryValueEx(hRegKey, "Mail Address",NULL,&datatype,(LPBYTE)szMailAddr,&datasz);
		datasz = 256;
		RegQueryValueEx(hRegKey, "Pop3 Address",NULL,&datatype,(LPBYTE)szPopServer,&datasz);
		datasz = 256;
		RegQueryValueEx(hRegKey, "User Name",NULL,&datatype,(LPBYTE)szUserName,&datasz);
		datasz = 256;
		RegQueryValueEx(hRegKey, "Password",NULL,&datatype,(LPBYTE)szPass,&datasz);
		datasz = 256;
		RegQueryValueEx(hRegKey, "From Name",NULL,&datatype,(LPBYTE)szName,&datasz);
		datasz = sizeof(DWORD);
		RegQueryValueEx(hRegKey, "APop",NULL,&datatype,(LPBYTE)&bApop,&datasz);
		datasz = sizeof(DWORD);
		RegQueryValueEx(hRegKey, "Pop Before SMTP",NULL,&datatype,(LPBYTE)&bPSmtp,&datasz);
		datasz = sizeof(DWORD);
		RegQueryValueEx(hRegKey, "Server Mail Del",NULL,&datatype,(LPBYTE)&bSMailDel,&datasz);
	}

	// レジストリを閉じる
	RegCloseKey(hRegKey);

}

// レジストリの保存
void SaveRegKey()
{
	HKEY hRegKey;			// レジストリハンドラ
	unsigned long Result;	// レジストリ用の変数
	
	// レジストリを開く
	RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\Angelic Software\\MailViewer", 0, "MailViewer", 0, KEY_ALL_ACCESS, NULL, &hRegKey, &Result);

	// レジストリへの書き込み
	RegSetValueEx(hRegKey, "Smtp Address", 0, REG_SZ, (LPBYTE)szSmtpServer, strlen(szSmtpServer)+1);
	RegSetValueEx(hRegKey, "Pop3 Address", 0, REG_SZ, (LPBYTE)szPopServer, strlen(szPopServer)+1);
	RegSetValueEx(hRegKey, "From Name", 0, REG_SZ, (LPBYTE)szName, strlen(szName)+1);
	RegSetValueEx(hRegKey, "Mail Address", 0, REG_SZ, (LPBYTE)szMailAddr, strlen(szMailAddr)+1);
	RegSetValueEx(hRegKey, "User Name", 0, REG_SZ, (LPBYTE)szUserName, strlen(szUserName)+1);
	RegSetValueEx(hRegKey, "Password", 0, REG_SZ, (LPBYTE)szPass, strlen(szPass)+1);
	RegSetValueEx(hRegKey, "APop", 0, REG_DWORD, (LPBYTE)&bApop, sizeof(DWORD));
	RegSetValueEx(hRegKey, "Pop Before SMTP", 0, REG_DWORD, (LPBYTE)&bPSmtp, sizeof(DWORD));
	RegSetValueEx(hRegKey, "Server Mail Del", 0, REG_DWORD, (LPBYTE)&bSMailDel, sizeof(DWORD));

	// レジストリを閉じる
	RegCloseKey(hRegKey);

}

// メール格納ディレクトリの初期化
BOOL InitFolder()
{
	LPSTR szFolder;
	
	szFolder = (char *)malloc(sizeof(char) * MAX_PATH);
	
	// Inboxフォルダの確認
	GetCurrentDirectory(MAX_PATH, szFolder);
	strcat(szFolder, "\\inbox");
	
	// Inboxフォルダがないとき
	if(FileExists(szFolder) != TRUE){
		// フォルダを作成する
		if(!CreateDirectory(szFolder, NULL)){
			free(szFolder);
			return FALSE;
		}
	}

	free(szFolder);

	return TRUE;
}

// メールの受信
BOOL RcvMail(HWND hWnd, HWND hList)
{
	// 基本設定がされていないとき
	if(strcmp(szPopServer, "") == 0 || strcmp(szUserName, "") == 0 || strcmp(szPass, "") == 0){
		MessageBox(hWnd, "メール受信の設定ができていません", "エラー", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// まずはPOP3 サーバーに接続
	if((s = NMailPop3Connect(szPopServer)) != INVALID_SOCKET){
		// 接続メッセージを追加
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"メールを受信しています");

		// POP3 認証 & メールタイトル一覧表示
		if(List_Mail(s, szPopServer, szUserName, szPass, hList, hStatus) >= 0){
			NMailPop3Close(s);
			return TRUE;
		}
		else{
			MessageBox(hWnd, "認証に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
			NMailPop3Close(s);
			return FALSE;
		}
	}
	else{
		MessageBox(hWnd, "接続に失敗しました", "エラー", MB_OK | MB_ICONSTOP);
		NMailPop3Close(s);
		return FALSE;
	}
}

int CALLBACK ListCompProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParam3)
{
	static LVFINDINFO lvf;
	static int nItem1, nItem2;
	static char buf1[MAX_PATH], buf2[MAX_PATH];
	SORTDATA *lpsd;

	lpsd = (SORTDATA *)lParam3;

	lvf.flags = LVFI_PARAM;
	lvf.lParam = lParam1;
	nItem1 = ListView_FindItem(lpsd->hwndList, -1, &lvf);

	lvf.lParam = lParam2;
	nItem2 = ListView_FindItem(lpsd->hwndList, -1, &lvf);

	ListView_GetItemText(lpsd->hwndList, nItem1, lpsd->isortSubItem, buf1, sizeof(buf1));
	ListView_GetItemText(lpsd->hwndList, nItem2, lpsd->isortSubItem, buf2, sizeof(buf2));

	if(lpsd->isortSubItem != 1){
		if(lpsd->iUPDOWN == UP){
			return(stricmp(buf1, buf2));
		}
		else{
			return(stricmp(buf1, buf2) * -1);
		}
	}
	else{
		if(lpsd->iUPDOWN == UP){
			if(atoi(buf1) > atoi(buf2)){
				return 1;
			}
			else if(atoi(buf1) == atoi(buf2)){
				return 0;
			}
			else{
				return -1;
			}
		}
		else{
			if(atoi(buf1) > atoi(buf2)){
				return -1;
			}
			else if(atoi(buf1) == atoi(buf2)){
				return 0;
			}
			else{
				return 1;
			}
		}
	}
}

// メッセージポンプ
void MessagePump()
{

	MSG msg;

	// PeekMessage関数を呼び出して
	if(PeekMessage(&msg, (HWND)NULL, 0, 0, PM_REMOVE)){
		// メッセージが届いていたら翻訳 & 配信する
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}

