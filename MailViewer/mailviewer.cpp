///////////////////////////////////////////////////////////////////////////////////////
// Mail Viewer Version 0.07
// �쐬��:���q����(Angelic Software)
// �X�V�o�[�W����:0.07(2007�N12��27��)
// http://www.�V�g�̂�����.com/
//////////////////////////////////////////////////////////////////////////////////////

#pragma warning(disable : 4996)

#include "mailviewer.h"

// WinMain�֐�
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInst, LPSTR lpszCmd, int nCmdShow){

	MSG msg;
	HWND hFind;
	HACCEL hAc;
	INITCOMMONCONTROLSEX ic;	// �R�����R���g���[��

	// ��d�N���֎~
	if((hFind = FindWindow(NULL, "Mail Viewer")) != NULL) {
		SetForegroundWindow(hFind);
		return(0);
	}

	// ���[���i�[�f�B���N�g���̏�����
	if(!InitFolder()){
		MessageBox(NULL, "���[���t�H���_�̍쐬�Ɏ��s���܂����B", "�G���[", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// nMail.dll�̏�����
	NMailSetParameter(0, TEMP_MAX, 0);
	NMailInitializeWinSock();
	
	// �R�����R���g���[���̏�����
	ic.dwSize = sizeof(INITCOMMONCONTROLSEX);
	ic.dwICC = ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&ic);

	// �A�v���P�[�V�����̏�����
	if(!InitApp(hInstance)){
		return FALSE;
	}

	// �O���[�o���ϐ��ɃR�s�[
	hInst = hInstance;

	// �E�B���h�E�̍쐬
	if(!InitInstance(hInstance, nCmdShow)){
		return FALSE;
	}

	// �A�N�Z�����[�^�L�[��HTML�w���v�̏�����
	hAc = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
	HtmlHelp(NULL, NULL, HH_INITIALIZE, (DWORD)&dwCookie);

	while(bRet = GetMessage(&msg, NULL, 0, 0) != 0){
		// GetMessage���G���[�̂Ƃ�
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

// �A�v���P�[�V�����̏�����
ATOM InitApp(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	// �E�B���h�E�N���X�̐ݒ�
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

// �E�B���h�E�̍쐬
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;

	// �E�B���h�E�̍쐬
	hWnd = CreateWindow(szWinName, "Mail Viewer", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);

	if(!hWnd){
		return FALSE;
	}

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
	
	// �O���[�o���ϐ��ɃR�s�[
	hParent = hWnd;

	return TRUE;
}

// �E�B���h�E�v���V�[�W��
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int Ret;
	static int iStatusWy;
	static HWND hEdit, hList;	// �R���g���[���n���h��
	static int sortsubno[NO_OF_SUBITEM] = {UP};
	static int isortsubno;

	RECT rc;					// RECT�\����
	POINT pt;					// POINT�\����
	DWORD dwStyle;				// �E�B���h�E�X�^�C��
	LPNMHDR lpnmhdr;			// �w�b�_�n���h��
	LPNMLISTVIEW lpnmlv;		// ���X�g�r���[�n���h��
	HMENU hMenu, hPMenu;		// ���j���[�n���h��
	SORTDATA SortData;			// ���X�g�̃\�[�g�p�\����
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

			// ���W�X�g���̓ǂݍ���
			LoadRegKey();

			// �N�����̃o�O�����
			No = -1;

			// ���[�J���ɕۑ����Ă��郁�[���ꗗ��\��
			SetLocalMail(hList);

			SortData.hwndList = hList;
			isortsubno = 2;
			SortData.isortSubItem = isortsubno;
			SortData.iUPDOWN = sortsubno[isortsubno];

			if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
				MessageBox(hWnd, "�\�[�g�Ɏ��s���܂����B", "�G���[", MB_OK | MB_ICONSTOP);
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

			// �V�����R�[�h
			if(lpnmhdr->hwndFrom == hList){
				switch(lpnmhdr->code){
					// �V���O���N���b�N�̎�
					case NM_CLICK:
						lpnmlv = (LPNMLISTVIEW)lParam;
						No = lpnmlv->iItem;
						
						// �N���b�N�����ԍ���-1�̎��̓��[����\�����Ȃ�
						if(No == -1){
							break;
						}
						
						// ���[���f�[�^�t�@�C�������擾����
						ListView_GetItemText(hList, No, 4, szSelMailFileName, sizeof(szSelMailFileName));

						// ���[���t�@�C�����J��
						Read_Mail(szSelMailFileName, hEdit);
					break;
					case NM_RCLICK:
						lpnmlv = (LPNMLISTVIEW)lParam;
						No = lpnmlv->iItem;
						
						// �N���b�N�����ԍ���-1�̎��̓��j���[��\�����Ȃ�
						if(No == -1){
							break;
						}

						// �E�N���b�N���j���[���J��
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
							MessageBox(hWnd, "�\�[�g�Ɏ��s���܂����B", "�G���[", MB_OK | MB_ICONSTOP);
						}

					break;
				}
			}
            break;
		// ���j���[��I���������̐����\��(�X�e�[�^�X�o�[)
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
		// ���j���[�𔲂������̏���
		case WM_EXITMENULOOP:
			SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
			break;
		case WM_COMMAND:
			switch(LOWORD(wParam)){
				case IDM_NEW:
					// �V�K
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG4), hWnd, (DLGPROC)NewMailProc);
					break;
				case IDM_OPEN:
					// �J��
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG5), hWnd, (DLGPROC)AttachFileProc);
					break;
				case IDM_SAVE:
					// �ۑ�
					if(No >= 0){
						WriteMailFile(hWnd);
					}
					break;
				case IDM_PRINT:
					// ���
					if(No >= 0){
						// ���[���̌������擾
						ListView_GetItemText(hList, No, 0, szSelMailSubject, sizeof(szSelMailSubject));
						wsprintf(strMsg, "�u%s�v��������܂����H", szSelMailSubject);

						Ret = MessageBox(hWnd, strMsg, "����̊m�F", MB_YESNO | MB_ICONQUESTION);
						if(Ret == IDYES){
							MailPrint(hWnd);
						}
					}
					break;
				case IDM_RCVMAIL:
					// ���[���̎�M
					// �G�f�B�b�g�{�b�N�X�̓��e���폜����
					SetWindowText(hEdit, "");

					// ���[����M�����̃T�u���[�`��
					RcvMail(hWnd, hList);

					SortData.hwndList = hList;
					SortData.isortSubItem = isortsubno;
					SortData.iUPDOWN = sortsubno[isortsubno];

					if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
						MessageBox(hWnd, "�\�[�g�Ɏ��s���܂����B", "�G���[", MB_OK | MB_ICONSTOP);
					}
					break;
				case IDM_REPLY:
					// ���[���̕ԐM
					if(No >= 0){
						bReply = TRUE;
						DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG4), hWnd, (DLGPROC)NewMailProc);
					}
					break;
				case IDM_DELETE:
					// ���[���̍폜
					if(No >= 0){
						// �e�L�X�g�{�b�N�X��������
						SetWindowText(hEdit, "");

						// ���[���f�[�^�t�@�C�������擾����
						ListView_GetItemText(hList, No, 4, szSelMailFileName, sizeof(szSelMailFileName));
						ListView_GetItemText(hList, No, 0, szSelMailSubject, sizeof(szSelMailSubject));

						wsprintf(strMsg, "�u%s�v���폜���܂����H", szSelMailSubject);
						Ret = MessageBox(hWnd, strMsg, "�폜�̊m�F", MB_YESNO | MB_ICONQUESTION);

						if(Ret == IDYES){
							// �I���������[�����폜
							if(!DeleteFile(szSelMailFileName)){
								// �폜�Ɏ��s�����Ƃ�
								MessageBox(hWnd, "�폜�Ɏ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
								break;
							}
							else{
								// �폜���o�����Ƃ�
								MessageBox(hWnd, "���[�����폜���܂���", "Mail Viewer", MB_OK | MB_ICONINFORMATION);
							}

							// ��M���[���ꗗ�̍X�V
							SetLocalMail(hList);

							SortData.hwndList = hList;
							SortData.isortSubItem = isortsubno;
							SortData.iUPDOWN = sortsubno[isortsubno];
							
							if(ListView_SortItems(hList, &ListCompProc, &SortData) != TRUE){
								MessageBox(hWnd, "�\�[�g�Ɏ��s���܂����B", "�G���[", MB_OK | MB_ICONSTOP);
							}
						}
					}
					break;
				case IDM_HEADER:
					// �w�b�_�\��
					if(No >= 0){
						DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG3), hWnd, (DLGPROC)HeaderProc);
						SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"");
					}
					break;
				case IDM_SET:
					// ���[���̐ݒ�
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
			// HTML�w���v���J���Ă��������
			if(HtmlHelp(hWnd, "mailviewer.chm", HH_GET_WIN_HANDLE, (DWORD)"mailviewer01")){
				HtmlHelp(NULL, NULL, HH_CLOSE_ALL, 0);
			}

			// ���W�X�g���̕ۑ�
			SaveRegKey();

			// �E�B���h�E�n���h����j��
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

// ���[���̐ݒ�_�C�A���O
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

					// APOP�̃`�F�b�N
					if(IsDlgButtonChecked(hDlg, IDC_CHECK1) == BST_CHECKED){
						bApop = TRUE;
					}
					else{
						bApop = FALSE;
					}

					// POP Before SMTP�̃`�F�b�N
					if(IsDlgButtonChecked(hDlg, IDC_CHECK2) == BST_CHECKED){
						bPSmtp = TRUE;
					}
					else{
						bPSmtp = FALSE;
					}

					// �T�[�o�Ƀ��[�����c���̃`�F�b�N
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

// �o�[�W�������_�C�A���O
LRESULT CALLBACK AboutProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hStatic;
	char strVerNo[256];
	double dNo;

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hStatic = GetDlgItem(hDlg, IDC_STATIC1);
			dNo = itod(VersionNo, 0.01);		// �����������_�t���̐��l�ɕϊ�
			sprintf(strVerNo, "%0.2f", dNo);	// wsprintf���Ɩ����Ȃ̂�sprintf�ɕύX
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

// �w�b�_���\���_�C�A���O
LRESULT CALLBACK HeaderProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;

	switch(msg){
		case WM_INITDIALOG:
			CenterWindow(hDlg, hParent);
			hEdit = GetDlgItem(hDlg, IDC_EDIT1);
			SetWindowText(hEdit, szHeaderMsg);
			wsprintf(strMsg, "\"%s\"�̃w�b�_���", szHeader);
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

// ���[���쐬�_�C�A���O
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
					MessageBox(hDlg, "�G���[���������܂����B", "�G���[", MB_OK | MB_ICONSTOP);
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
					// ���e�`�F�b�N(����A������)�@�\
					Edit_GetText(hAddr, szToData, sizeof(szToData));
					
					if(strcmp(szToData, "") == 0 || strstr(szToData, "@") == 0 || strstr(szToData, ".") == 0){
						MessageBox(hDlg, "���悪���͂���Ă��܂��񂩁A�s���ȃ��[���A�h���X�ł�", "�G���[", MB_OK | MB_ICONSTOP);
						return FALSE;
					}

					Edit_GetText(hSubject, szSubjectData, sizeof(szSubjectData));

					// ���������L�����M�s��
					if(strcmp(szSubjectData, "���������L��") == 0 || strcmp(szSubjectData, "�L��") == 0 || strcmp(szSubjectData, "����") == 0){
						MessageBox(hDlg, "���̃\�t�g�͍L�����[���𑗐M�ł��Ȃ����Ă��܂�", "�G���[", MB_OK | MB_ICONSTOP);
						return FALSE;
					}

					if(strcmp(szSubjectData, "") == 0){
						NoTitle = MessageBox(hDlg, "���������͂���Ă��܂��񂪑��M���Ă�낵���ł����H\n�͂��̏ꍇ�͌�����(����)�ő��M����܂�", "����������", MB_YESNO | MB_ICONINFORMATION);
						if(NoTitle == IDNO){
							return FALSE;
						}
						else{
							strcpy(szSubjectData, "(����)");
						}
					}
					
					Edit_GetText(hEdit, szBodyData, sizeof(szBodyData));
					Edit_GetText(hAttach, szFileData, sizeof(szFileData));

					// Pop Before SMTP���I������Ă��鎞��POP3�ɐڑ����Ă��瑗�M����
					if(bPSmtp == TRUE){
						// POP3�T�[�o�ɐڑ����Đؒf����
						bResult = SmtpPop(s, szPopServer, szUserName, szPass);

						// �߂�l��FALSE�̎��͐ڑ����s
						if(bResult == FALSE){
							// POP3�T�[�o�ڑ����s
							MessageBox(hDlg, "POP3�T�[�o�ɐڑ��ł��܂���ł����B\nPOP3�T�[�o�̐ݒ���Ċm�F���Ă��������B", "�G���[", MB_OK | MB_ICONSTOP);
							return FALSE;
						}
					}

					// ���[���𑗐M����
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

// �Y�t�t�@�C����W�J���ĕۑ��_�C�A���O
LRESULT CALLBACK AttachFileProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	HWND hList, hFDEdit, hFBtn, hFDBtn;
	
	BROWSEINFO bi;
	ITEMIDLIST *lpid;
	HRESULT hr;
	LPMALLOC pMalloc = NULL;				// IMalloc�ւ̃|�C���^
	BOOL bDir = FALSE;

	HANDLE hFile;							// �t�@�C���n���h��
	HGLOBAL hMem;							// �O���[�o���n���h��
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// �o�C�g��
	char *lpszBuf;							// �t�@�C���o�b�t�@
	char f_id[TEMP_MAX], *PertId;			// �������[����ID
	int no;									// �Y�t�t�@�C���̌�
	int ret;								// �߂�l

	// �Y�t���[���W�J�p�ϐ�
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

				// �t�@�C���������͂���Ă��Ȃ��Ƃ�
				if(strlen(szAtFileName) == 0){
					return FALSE;
				}

				// �t�@�C�����J��
				hFile = CreateFile(szAtFile, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
				
				if(hFile == INVALID_HANDLE_VALUE){
					MessageBox(hDlg, "�t�@�C�����I�[�v���ł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
					return FALSE;
				}
				
				// �t�@�C���T�C�Y�𒲂ׂ�
				dwFSize = GetFileSize(hFile, &dwFSizeHigh);
				
				if(dwFSizeHigh != 0){
					MessageBox(hDlg, "�t�@�C�����傫�����܂�", "�G���[", MB_OK | MB_ICONSTOP);
					CloseHandle(hFile);
					return FALSE;
				}

				// �t�@�C����ǂݍ��ނ��߂̃������̈���m��
				hMem = GlobalAlloc(GHND, dwFSize + 1);
				
				if(hMem == NULL){
					MessageBox(hDlg, "���������m�ۂł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
					CloseHandle(hFile);
					return FALSE;
				}
				
				// �������̈�����b�N���ăt�@�C����ǂݍ���
				lpszBuf = (char *)GlobalLock(hMem);
				ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
				lpszBuf[dwFSize] = '\0';
				
				// szAtFileData�̃������̈���m�ۂ���
				szAtFileData = (char *)malloc(sizeof(char) * dwFSize);

				if(szAtFileData == NULL){
					MessageBox(hDlg, "�������m�ۂɎ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
					return FALSE;
				}

				// ���[���t�@�C���̒��g���R�s�[����
				strcpy((char *)szAtFileData, lpszBuf);
				
				// �t�@�C�������
				CloseHandle(hFile);
				GlobalUnlock(hMem);
				GlobalFree(hMem);
				
				// �t�@�C�����Y�t����Ă��邩���m�F����
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
					// �Y�t�t�@�C�����Ȃ��̂Ń��b�Z�[�W���o���ďI��
					MessageBox(hDlg, "���̃��[���ɁA�Y�t�t�@�C���͂���܂���", "�Y�t�t�@�C���Ȃ�", MB_OK |MB_ICONINFORMATION);
				}

				return TRUE;

			case IDC_FILEBTN:
				memset(&bi, 0, sizeof(BROWSEINFO));
				bi.hwndOwner = hDlg;
				bi.lpfn = SHMyProc;
				bi.ulFlags = BIF_EDITBOX | BIF_STATUSTEXT | BIF_VALIDATE;
				bi.lpszTitle = "�W�J��f�B���N�g���w��";
				lpid = SHBrowseForFolder(&bi);
				
				if(lpid == NULL){
					return FALSE;
				}
				else{
					hr = SHGetMalloc(&pMalloc);
					if(hr == E_FAIL){
						MessageBox(hDlg, "SHGetMalloc Error", "�G���[", MB_OK | MB_ICONSTOP);
						return FALSE;
					}
					SHGetPathFromIDList(lpid, szDir);
					if(szDir[strlen(szDir) - 1] != '\\'){
						strcat(szDir, "\\");
					}

					hFDEdit = GetDlgItem(hDlg, IDC_EDIT1);

					// �W�J��t�H���_���Z�b�g����
					SetWindowText(hFDEdit, szDir);

					pMalloc->Free(lpid);
					pMalloc->Release();
					bDir = TRUE;
				}
				return TRUE;

				case IDOK:
					// �Y�t�t�@�C���̓W�J
					ret = NMailAttachmentFileFirst(Temp, subject, date, from, header, body, szDir, FileName, szAtFileData, NULL);
					if(ret == NMAIL_SUCCESS){
						NMailAttachmentFileClose(Temp);
						wsprintf(strMsg, "%s��%s��W�J���܂����B", szDir, FileName);
						MessageBox(hDlg, strMsg, "Mail Viewer", MB_OK | MB_ICONINFORMATION);
					}
					// �������̉��
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

// ���X�g�Ɍ����A���o�l�A���t�̃J������}������
void InsertColumn(HWND hList)
{
    LVCOLUMN lvcol;

    lvcol.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvcol.fmt = LVCFMT_LEFT;
    lvcol.cx = 300;
    lvcol.pszText = "����";
    lvcol.iSubItem = 0;
    ListView_InsertColumn(hList, 0, &lvcol);

    lvcol.cx = 250;
    lvcol.pszText = "���o�l";
    lvcol.iSubItem = 1;
    ListView_InsertColumn(hList, 1, &lvcol);

    lvcol.cx = 155;
    lvcol.pszText = "���t";
    lvcol.iSubItem = 2;
    ListView_InsertColumn(hList, 2, &lvcol);

    lvcol.cx = 50;
    lvcol.pszText = "�T�C�Y";
    lvcol.iSubItem = 3;
    ListView_InsertColumn(hList, 3, &lvcol);

	// ���[���̃t�@�C���p�X���i�[(��\��)
    lvcol.cx = 0;
    lvcol.pszText = "";
    lvcol.iSubItem = 4;
    ListView_InsertColumn(hList, 4, &lvcol);

    return;
}

// ���X�g�Ɍ����A���o�l�A���t��}������
void InsertItem(HWND hList, int iItemNo, int iSubNo, char *lpszText)
{
	LVITEM item;
	int lvitem;

	memset(&item, 0, sizeof(LVITEM));
	lvitem = iItemNo - 1;	// lvitem�̏����l��iItemNo - 1�̒l����

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

// ���[���̃��X�g��\������
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

	// �F��
	if((max = NMailPop3Authenticate(s, id, pass, bApop)) >= 0){
		// �T�[�o�ɂ��郁�[���̐���0�ȏ�̂Ƃ�
		if(max > 0 ){
			// �v���O���X�p�[�\��
			hProgress = CreateWindowEx(0, PROGRESS_CLASS, "", WS_CHILD | WS_VISIBLE, sb_size, 5, 100, 16, hStatus, (HMENU)ID_PROGRESS, hInst, NULL);
			SendMessage(hProgress, PBM_SETRANGE, 0, MAKELONG(0, max));
		}
		else{
			// �v���O���X�o�[���폜
			DestroyWindow(hProgress);
		}

		// ��M���Ă��郁�[���̐����w�b�_��ǂݏo��
		for(count = 0; count < max; count++){

			// �v���O���X�o�[�𑝉�
			SendMessage(hProgress, PBM_SETPOS, count + 1, 0);

			if(NMailPop3GetMailStatus(s, count + 1, subject, date, from, header, FALSE) >= 0){

				// ���܂������Ȃ������̂ŃT�C�Y�擾���@��ύX
				if((msize = NMailPop3GetMailSize(s, count + 1)) >= 0){
					sprintf(szSize, "%d", msize);
				}
				else{
					wsprintf(szSize, "0");
				}

				// �擾�������t�𕪉�
				mid(date, szlYear, 1, 4);
				mid(date, szlMonth, 6, 2);
				mid(date, szlDay, 9, 2);
				mid(date, szlHour, 12, 2);
				mid(date, szlMinute, 15, 2);
				mid(date, szlSecond, 18, 2);

				// UIDL�̃T�C�Y���m��
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

				// �J�����g�f�B���N�g�����擾����
				GetCurrentDirectory(MAX_PATH, lpFilePath);
				strcat(lpFilePath, "\\inbox\\");

				// UIDL���擾����
				NMailPop3GetUidl(s, count + 1, lpUidl, 1024);

				// ���X�g�p�̃t�@�C�������쐬
				wsprintf(szName, "%s%s%s%s%s%s_%s.dat", szlYear, szlMonth, szlDay, szlHour, szlMinute, szlSecond, lpUidl);

				strcat(lpFilePath, szName);
				bFile = FileExists(lpFilePath);

				// �t�@�C�������݂��Ȃ�(�V�K���[��)�Ƃ�
				if(bFile == FALSE){
					// �������̃��b�Z�[�W�̕\��
					wsprintf(strMsg, "%d���ڂ�������", count + 1);
					SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)strMsg);

					// ���[���f�[�^��ۑ�
					Store_Mail(s, count + 1, lpFilePath);
				}
				else{
					// ��M�ς݂̏ꍇ�͎�M�ς݃J�E���g�𑝉�����
					rcount++;
				}

				// �T�[�o�Ƀ��[�����c���`�F�b�N���͂�����Ă���Ƃ�
				if(bSMailDel == FALSE){
					// �T�[�o�̃��[�����폜����
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
		// �F�؃G���[�̏ꍇ(max�ɂ̓G���[�l���i�[����Ă���)
		return max;
	}

	if(max > 0){
		// �v���O���X�o�[���폜
		DestroyWindow(hProgress);
	}

	// �V�K���[�������M�ς݃��[���̐�������
	newcnt = max - rcount;

	// �V�����[��������Ƃ�
	if(newcnt > 0){
		// �ۑ����ꂽ���[���̈ꗗ��\������
		SetLocalMail(hList);
	}

	// �X�e�[�^�X���b�Z�[�W�̏�����
	strcpy(strMsg, "");

	// �ő匏����0��(���[�����Ȃ�)�̏ꍇ
	if(newcnt == 0){
		// ���[�����T�[�o�ɂȂ�(�܂��͊��ǃ��[�������Ȃ�)��
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"���[���͓͂��Ă��܂���");
	}
	else{
		// ���[��������ꍇ�͌����\��
		wsprintf(strMsg, "%d���̃��[�����͂��Ă��܂�", newcnt);
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)strMsg);
	}

	return max;
}

//////////////////////////////////////////////////////////////////////////
// Send_Mail
// ���e�@�F���[���𑗐M����
// �߂�l�F�Ȃ�
//////////////////////////////////////////////////////////////////////////
void Send_Mail(char *host, char *to, char *subject, char *body, char *atatch)
{
	char from[SEND_MAX];

	// from�̕����̃t�H�[�}�b�g�쐬
	sprintf(from, "%s <%s>", szName, szMailAddr);

	// �{���̏I����\r\n��t����
	strcat(body, "\r\n");

    // ���[�����M
	if(NMailSmtpSendMail(host, to, NULL, NULL, from, subject, body, "X-MAILER: Mail Viewer Version 0.0.7", atatch, 0) >= 0){
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"���[���𑗐M���܂���");
	}
	else{
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"���[�����M�Ɏ��s���܂���");
	}
}

// �_�C�A���O�𒆉��ɕ\������
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

// �Y�t�t�@�C���̏���
void SelectAttachFile(HWND hDlg, char *lpstrFile, char *lpstrTitle)
{
	OPENFILENAME ofn;

	memset(&ofn, 0, sizeof(OPENFILENAME));

	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hDlg;
	ofn.lpstrFilter = "���ׂẴt�@�C��(*.*)\0*.*\0\0";
	ofn.lpstrFile = lpstrFile;
	ofn.lpstrFileTitle = lpstrTitle;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = "txt";
	ofn.lpstrTitle = "�Y�t�t�@�C���̑I��";
	ofn.nMaxFileTitle = MAX_PATH;
	GetOpenFileName(&ofn);

	return;
}

// ���[���t�@�C�����J��
void AttachFileOpen(HWND hDlg, char *lpstrFile, char *lpstrTitle)
{
	OPENFILENAME ofn;	// �J���_�C�A���O�\����

	memset(&ofn, 0, sizeof(OPENFILENAME));

	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hDlg;
	ofn.lpstrFilter = "Mail Viewer(*.mvr)\0*.mvr\0���[���t�@�C��(*.eml)\0*.eml\0�e�L�X�g�t�@�C��(*.txt)\0*.txt\0���ׂẴt�@�C��(*.*)\0*.*\0";
	ofn.lpstrFile = lpstrFile;
	ofn.lpstrFileTitle = lpstrTitle;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = "mvr";
	ofn.lpstrTitle = "���[���t�@�C�����J��";
	ofn.nMaxFileTitle = MAX_PATH;
	GetOpenFileName(&ofn);

	return;
}

// ���[�����e�L�X�g�ŏo��
int WriteMailFile(HWND hWnd)
{
	OPENFILENAME ofn;
	HANDLE hFile;
	DWORD dwAccBytes;

	memset(&ofn, 0, sizeof(OPENFILENAME));
	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.hwndOwner = hWnd;
	ofn.lpstrFilter = "Mail Viewer(*.mvr)\0*.mvr\0���[���t�@�C��(*.eml)\0*.eml\0�e�L�X�g�t�@�C��(*.txt)\0*.txt\0���ׂẴt�@�C��(*.*)\0*.*\0";
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
		MessageBox(hWnd, "�t�@�C����ۑ��ł��܂���ł���", "�G���[", MB_OK | MB_ICONSTOP);
	}
	return 0;
}

//////////////////////////////////////////////////////////////////////////
// itod(int to double)
// ���e�@�F�����������_�t���̐��l�ɕϊ�
// �����@�Fnum�@ �E�E�E ����(123�̂悤�ɋL�q/int)
// �@�@�@�@point �E�E�E �����_�̈ʒu(0.01�ȂǂŎw��/double)
// �߂�l�F�����_�������l(double)�A-1���Ԃ�ꍇ�̓G���[
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
// ���e�@�F���[���̓��e���������(�ȈՈ��)
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
	sprintf(szPrnDat, "\r\n���M�����F%s \r\n���M�ҁF%s \r\n�����F%s \r\n ------------------------------------------------------------ \r\n %s", RepDate, RepFrom, RepSubject, strText);
	rct.left   = 10;
	rct.right  = tm.tmHeight * strlen(szPrnDat);
	rct.bottom = 10 + (tm.tmHeight * strlen(szPrnDat));
	DrawText(pd.hDC, szPrnDat, -1, &rct, DT_LEFT);

	EndPage(pd.hDC);
	EndDoc(pd.hDC);
	DeleteDC(pd.hDC);

	return(0);

}

// �Y�t�t�@�C���i�[��t�H���_��I������
int CALLBACK SHMyProc(HWND hWnd, UINT msg, LPARAM lp1, LPARAM lp2)
{
    char szFolder[MAX_PATH];
    char szBuf[1024];
    ITEMIDLIST *lpid;

    switch (msg) {
        case BFFM_SELCHANGED:
            lpid = (ITEMIDLIST *)lp1;
            SHGetPathFromIDList(lpid, szFolder);
            wsprintf(szBuf, "���݁u%s�v���I������Ă��܂�", szFolder);
            SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, (LPARAM)szBuf);
            return 0;
        case BFFM_VALIDATEFAILED:
            MessageBox(hWnd, "�s���ȃt�H���_�ł�", "�G���[", MB_OK | MB_ICONSTOP);
            return 1;

    }
    return 0;
}

// POP Before SMTP�p��POP3�F��
BOOL SmtpPop(SOCKET s, char *host, char *id, char *pass)
{
	BOOL bSuccess;

	// POP3�T�[�o�ɐڑ�
	if((s = NMailPop3Connect(host)) != INVALID_SOCKET){
		// POP3�T�[�o�̔F��
		if(NMailPop3Authenticate(s, id, pass, bApop) >= 0){
			// 0�ȏ�Ȃ�TRUE��Ԃ�
			bSuccess = TRUE;
		}
		else{
			// -1�ȉ��Ȃ�FALSE��Ԃ�
			bSuccess = FALSE;
		}
		// POP3�T�[�o���甲����
		NMailPop3Close(s);
	}
	else{
		// INVALID_SOCKET�̎���FALSE��Ԃ�
		bSuccess = FALSE;
	}
	return bSuccess;
}

// ���[���f�[�^��ۑ�����
int SaveMailData(LPSTR lpFilePath)
{
	HANDLE hFile;
	DWORD dwAccBytes;
	hFile = CreateFile(lpFilePath, GENERIC_WRITE, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	
	// �t�@�C�����쐬�ł��Ȃ������Ƃ�
	if(hFile == INVALID_HANDLE_VALUE){
		return -1;
	}

	WriteFile(hFile, szMailData, strlen(szMailData) + 1, &dwAccBytes, NULL);
	
	if(CloseHandle(hFile) == 0){
		MessageBox(NULL, "�t�@�C����ۑ��ł��܂���ł���", "�G���[", MB_OK | MB_ICONSTOP);
		return -1;
	}
	return 0;
}

// List_Mail�̎��Ƀ��[���f�[�^��ۑ�����
void Store_Mail(SOCKET s, int no, LPSTR lpFilePath/*lpUidl*/)
{
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX];
	int  size;

	// ���[���̃T�C�Y���m�ۂ��ăt�@�C���ɏ�������
	if((size = NMailPop3GetMailSize(s, no)) >= 0){
		if((body =(char *)malloc(size)) != NULL){
			if(NMailPop3GetMail(s, no, subject, date, from, header, body, NULL, NULL) >= 0){
				sprintf(szMailData, "%s\r\n%s", header, body);
				SaveMailData(lpFilePath);
			}
			else{
				MessageBox(NULL, "�G���[���������܂���", "�G���[", MB_OK | MB_ICONSTOP);
			}
			free(body);
		}
	}
}

// MID�֐�
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

// ���[�J���ɕۑ�����Ă��郁�[�����擾����
void SetLocalMail(HWND hListL)
{
	WIN32_FIND_DATA wfd;
	HANDLE hFind;
	HANDLE hFile;							// �t�@�C���n���h��
	HGLOBAL hMem;							// �O���[�o���n���h��
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// �o�C�g��
	char szfname[MAX_PATH + 5];				// �t�@�C����
	char *lpszBuf;							// �t�@�C���o�b�t�@
	int iListNo;

	// �Y�t���[���W�J�p�ϐ�
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX], body[TEMP_MAX];
	char FileName[TEMP_MAX], Temp[NMAIL_ATTACHMENT_TEMP_SIZE], szSize[TEMP_MAX];

	// ListView�̕`����~�߂�
	SetWindowRedraw(GetDesktopWindow(), FALSE);

	// ���X�g�̏�����
	ListView_DeleteAllItems(hListL);
	GetCurrentDirectory(MAX_PATH, szfname);
	strcat(szfname, "\\inbox\\*.dat");
	iListNo = 0;

	// ���[���t�@�C����T��
	hFind = FindFirstFile(szfname, &wfd);

	// ���������s�����ꍇ(�t�@�C�����Ȃ��ꍇ)
	if(hFind == INVALID_HANDLE_VALUE){
		FindClose(hFind);
		return;
	}
	
	// �t�@�C�����Ȃ��Ȃ�܂Ń��X�g�ɒǉ�
	do{
		// �J�����g�f�B���N�g�����擾����
		GetCurrentDirectory(MAX_PATH, szfname);
		strcat(szfname, "\\inbox\\");

		// �����t�@�C�������R�s�[
		strcat(szfname, wfd.cFileName);

		// �t�@�C�����J��
		hFile = CreateFile(szfname, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		
		if(hFile == INVALID_HANDLE_VALUE){
			MessageBox(NULL, "�t�@�C�����I�[�v���ł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
			return;
		}
		
		// �t�@�C���T�C�Y�𒲂ׂ�
		dwFSize = GetFileSize(hFile, &dwFSizeHigh);
		
		if(dwFSizeHigh != 0){
			MessageBox(NULL, "�t�@�C�����傫�����܂�", "�G���[", MB_OK | MB_ICONSTOP);
			CloseHandle(hFile);
			return;
		}
		
		// �t�@�C����ǂݍ��ނ��߂̃������̈���m��
		hMem = GlobalAlloc(GHND, dwFSize + 1);
		
		if(hMem == NULL){
			MessageBox(NULL, "���������m�ۂł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
			CloseHandle(hFile);
			return;
		}
		
		// �������̈�����b�N���ăt�@�C����ǂݍ���
		lpszBuf = (char *)GlobalLock(hMem);
		ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
		lpszBuf[dwFSize] = '\0';
		
		// szAtFileData�̃������̈���m�ۂ���
		szAtFileData = (char *)malloc(sizeof(char) * dwFSize);
		
		if(szAtFileData == NULL){
			MessageBox(NULL, "�������m�ۂɎ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
			return;
		}
		
		// ���[���t�@�C���̒��g���R�s�[����
		strcpy((char *)szAtFileData, lpszBuf);
		
		// �t�@�C�������
		CloseHandle(hFile);
		GlobalUnlock(hMem);
		GlobalFree(hMem);
		
		// �����A��M���t�A���o�l�̏�����
		strcpy(subject, "");
		strcpy(date, "");
		strcpy(from, "");

		// �Y�t�t�@�C���̓W�J(���ۂ̓��[���{���ƃw�b�_�̕����̂�)
		NMailAttachmentFileFirst(Temp, subject, date, from, header, body, NULL, FileName, szAtFileData, NULL);

		// ���[���T�C�Y��ϊ�����
		wsprintf(szSize, "%d", dwFSize);

		// �������󔒂̎���(�����Ȃ�)��\������
		if(strlen(subject) > 0){
			InsertItem(hListL, iListNo, 0, subject);
		}
		else{
			InsertItem(hListL, iListNo, 0, "(�����Ȃ�)");
		}
		InsertItem(hListL, iListNo + 1, 1, from);
		InsertItem(hListL, iListNo + 1, 2, date);
		InsertItem(hListL, iListNo + 1, 3, szSize);
		InsertItem(hListL, iListNo + 1, 4, szfname);

		free(szAtFileData);

		iListNo++;

	}while(FindNextFile(hFind, &wfd) == TRUE);

	FindClose(hFind);

	// ListView�̕`����ĊJ����
	SetWindowRedraw(GetDesktopWindow(), TRUE);

	return;

}

// ���X�g�őI�����������ɑΉ����郁�[���f�[�^��ǂݏo��
void Read_Mail(char *szFName, HWND hEdit)
{
	HANDLE hFile;							// �t�@�C���n���h��
	HGLOBAL hMem;							// �O���[�o���n���h��
	DWORD dwFSizeHigh, dwFSize, dwAccBytes;	// �o�C�g��
	char *lpszBuf;							// �t�@�C���o�b�t�@
	
	// �Y�t���[���W�J�p�ϐ�
	char subject[TEMP_MAX], date[TEMP_MAX], from[TEMP_MAX], header[TEMP_MAX], body[TEMP_MAX];
	char FileName[TEMP_MAX], Temp[NMAIL_ATTACHMENT_TEMP_SIZE];
	
	// �ԐM�p�ϐ��̏�����
	strcpy(RepSubject, "");			// �ԐM�p�̌����쐬�p�ϐ��̏�����
	strcpy(RepFrom, "");			// �ԐM�p�̈���p�ϐ��̏�����
	strcpy(RepDate, "");			// �ԐM�p�̓��t�p�ϐ��̏�����
	strcpy(RepSubject, subject);	// �ԐM�p�̌����쐬�p�ϐ�
	strcpy(RepFrom, from);			// �ԐM�p�̈���p�ϐ�
	strcpy(RepDate, date);			// �ԐM�p�̓��t�p�ϐ�

	// �t�@�C�����J��
	hFile = CreateFile(szFName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	
	if(hFile == INVALID_HANDLE_VALUE){
		MessageBox(NULL, "�t�@�C�����I�[�v���ł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
		return;
	}
	
	// �t�@�C���T�C�Y�𒲂ׂ�
	dwFSize = GetFileSize(hFile, &dwFSizeHigh);
	
	if(dwFSizeHigh != 0){
		MessageBox(NULL, "�t�@�C�����傫�����܂�", "�G���[", MB_OK | MB_ICONSTOP);
		CloseHandle(hFile);
		return;
	}
	
	// �t�@�C����ǂݍ��ނ��߂̃������̈���m��
	hMem = GlobalAlloc(GHND, dwFSize + 1);
	
	if(hMem == NULL){
		MessageBox(NULL, "���������m�ۂł��܂���", "�G���[", MB_OK | MB_ICONSTOP);
		CloseHandle(hFile);
		return;
	}
	
	// �������̈�����b�N���ăt�@�C����ǂݍ���
	lpszBuf = (char *)GlobalLock(hMem);
	ReadFile(hFile, lpszBuf, dwFSize, &dwAccBytes, NULL);
	lpszBuf[dwFSize] = '\0';
	
	// szAtFileData�̃������̈���m�ۂ���
	szAtFileData = (char *)malloc(sizeof(char) * dwFSize);
	
	if(szAtFileData == NULL){
		MessageBox(NULL, "�������m�ۂɎ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
		return;
	}
	
	// ���[���t�@�C���̒��g���R�s�[����
	strcpy((char *)szAtFileData, lpszBuf);
	
	// �t�@�C�������
	CloseHandle(hFile);
	GlobalUnlock(hMem);
	GlobalFree(hMem);
	
	// �����A��M���t�A���o�l�̏�����
	strcpy(subject, "");
	strcpy(date, "");
	strcpy(from, "");

	// �Y�t�t�@�C���̓W�J(���ۂ̓��[���{���ƃw�b�_�̕����̂�)
	NMailAttachmentFileFirst(Temp, subject, date, from, header, body, NULL, FileName, szAtFileData, NULL);

	// �ǂݍ��񂾖{����\��
	sprintf(strText, "%s", body);
	strcpy(szMailData, szAtFileData);
	SetWindowText(hEdit, strText);
	strcpy(szHeaderMsg, header);

	// ���������݂���Ƃ�
	if(strlen(subject) > 0){
		strcpy(szHeader, subject);
	}
	else{
		strcpy(szHeader, "(�����Ȃ�)");
	}

}

// �t�@�C���̑��݊m�F
BOOL FileExists(LPSTR lpFileName)
{
	HANDLE hFile;			// �t�@�C���n���h��
	BOOL bFile;				// �t�@�C���t���O
	WIN32_FIND_DATA	wfd;	// WIN32_FIND_DATA�\����

	// �t�@�C���t���O��������
	bFile = FALSE;

	// �t�@�C������������
	hFile = FindFirstFile(lpFileName, &wfd);

	// �t�@�C�������݂���Ƃ�
	if(hFile != INVALID_HANDLE_VALUE){
		// �t�@�C���t���O��TRUE�ɂ���
		bFile = TRUE;
	}

	// �t�@�C���̌��������
	FindClose(hFile);

	return bFile;
}

// ���W�X�g���̓ǂݍ���
void LoadRegKey()
{
	HKEY hRegKey;							// ���W�X�g���n���h��
	unsigned long datasz, Result, datatype;	// ���W�X�g���p�̕ϐ�

	// ���W�X�g�����J��
	RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\Angelic Software\\MailViewer", 0, "MailViewer", 0, KEY_ALL_ACCESS, NULL, &hRegKey, &Result);

	// ���߂ă��W�X�g���ɓo�^����Ƃ�
	if(Result == REG_CREATED_NEW_KEY){
		// ���W�X�g���ւ̐V�K�쐬����
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
		// ���W�X�g������̓Ǐo������
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

	// ���W�X�g�������
	RegCloseKey(hRegKey);

}

// ���W�X�g���̕ۑ�
void SaveRegKey()
{
	HKEY hRegKey;			// ���W�X�g���n���h��
	unsigned long Result;	// ���W�X�g���p�̕ϐ�
	
	// ���W�X�g�����J��
	RegCreateKeyEx(HKEY_CURRENT_USER, "Software\\Angelic Software\\MailViewer", 0, "MailViewer", 0, KEY_ALL_ACCESS, NULL, &hRegKey, &Result);

	// ���W�X�g���ւ̏�������
	RegSetValueEx(hRegKey, "Smtp Address", 0, REG_SZ, (LPBYTE)szSmtpServer, strlen(szSmtpServer)+1);
	RegSetValueEx(hRegKey, "Pop3 Address", 0, REG_SZ, (LPBYTE)szPopServer, strlen(szPopServer)+1);
	RegSetValueEx(hRegKey, "From Name", 0, REG_SZ, (LPBYTE)szName, strlen(szName)+1);
	RegSetValueEx(hRegKey, "Mail Address", 0, REG_SZ, (LPBYTE)szMailAddr, strlen(szMailAddr)+1);
	RegSetValueEx(hRegKey, "User Name", 0, REG_SZ, (LPBYTE)szUserName, strlen(szUserName)+1);
	RegSetValueEx(hRegKey, "Password", 0, REG_SZ, (LPBYTE)szPass, strlen(szPass)+1);
	RegSetValueEx(hRegKey, "APop", 0, REG_DWORD, (LPBYTE)&bApop, sizeof(DWORD));
	RegSetValueEx(hRegKey, "Pop Before SMTP", 0, REG_DWORD, (LPBYTE)&bPSmtp, sizeof(DWORD));
	RegSetValueEx(hRegKey, "Server Mail Del", 0, REG_DWORD, (LPBYTE)&bSMailDel, sizeof(DWORD));

	// ���W�X�g�������
	RegCloseKey(hRegKey);

}

// ���[���i�[�f�B���N�g���̏�����
BOOL InitFolder()
{
	LPSTR szFolder;
	
	szFolder = (char *)malloc(sizeof(char) * MAX_PATH);
	
	// Inbox�t�H���_�̊m�F
	GetCurrentDirectory(MAX_PATH, szFolder);
	strcat(szFolder, "\\inbox");
	
	// Inbox�t�H���_���Ȃ��Ƃ�
	if(FileExists(szFolder) != TRUE){
		// �t�H���_���쐬����
		if(!CreateDirectory(szFolder, NULL)){
			free(szFolder);
			return FALSE;
		}
	}

	free(szFolder);

	return TRUE;
}

// ���[���̎�M
BOOL RcvMail(HWND hWnd, HWND hList)
{
	// ��{�ݒ肪����Ă��Ȃ��Ƃ�
	if(strcmp(szPopServer, "") == 0 || strcmp(szUserName, "") == 0 || strcmp(szPass, "") == 0){
		MessageBox(hWnd, "���[����M�̐ݒ肪�ł��Ă��܂���", "�G���[", MB_OK | MB_ICONSTOP);
		return FALSE;
	}
	
	// �܂���POP3 �T�[�o�[�ɐڑ�
	if((s = NMailPop3Connect(szPopServer)) != INVALID_SOCKET){
		// �ڑ����b�Z�[�W��ǉ�
		SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"���[������M���Ă��܂�");

		// POP3 �F�� & ���[���^�C�g���ꗗ�\��
		if(List_Mail(s, szPopServer, szUserName, szPass, hList, hStatus) >= 0){
			NMailPop3Close(s);
			return TRUE;
		}
		else{
			MessageBox(hWnd, "�F�؂Ɏ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
			NMailPop3Close(s);
			return FALSE;
		}
	}
	else{
		MessageBox(hWnd, "�ڑ��Ɏ��s���܂���", "�G���[", MB_OK | MB_ICONSTOP);
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

// ���b�Z�[�W�|���v
void MessagePump()
{

	MSG msg;

	// PeekMessage�֐����Ăяo����
	if(PeekMessage(&msg, (HWND)NULL, 0, 0, PM_REMOVE)){
		// ���b�Z�[�W���͂��Ă�����|�� & �z�M����
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}

