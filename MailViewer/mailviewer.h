///////////////////////////////////////////////////////////////////////////////////////
// Mail Viewer Version 0.07
// 作成者:砂倉瑞希(Angelic Software)
// 更新バージョン:0.07(2007年12月27日)
// http://www.天使のお茶会.com/
//////////////////////////////////////////////////////////////////////////////////////

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <winnetwk.h>
#include <shlobj.h>
#include <stdio.h>
#include <string.h>
#include <htmlhelp.h>
#include "nmail.h"
#include "resource.h"

#define ID_LISTVIEW 100
#define ID_STATUS   101
#define ID_EDIT     102
#define ID_PROGRESS 103

#define TEMP_MAX	65536		// 8192
#define SEND_MAX	94208		// 65536
#define MAIL_MAX    1024000

#define UP          1			// 昇順
#define DOWN        2			// 降順
#define NO_OF_SUBITEM 5			// サブアイテムの数

// ウィンドウプロシージャの型宣言
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK NewMailProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MailSetProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK HeaderProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK AboutProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK AttachFileProc(HWND, UINT, WPARAM, LPARAM);

// フォルダ選択ウィンドウプロシージャの型宣言
int CALLBACK SHMyProc(HWND, UINT, LPARAM, LPARAM);

// リスト比較用のプロシージャの型宣言
int CALLBACK ListCompProc(LPARAM, LPARAM, LPARAM);

// 関数のプロトタイプ宣言
int WriteMailFile(HWND);
double itod(int, double);
void InsertColumn(HWND);
void InsertItem(HWND, int, int, char *);
int List_Mail(SOCKET, char *, char *, char *, HWND, HWND);
void Send_Mail(char *, char *, char *, char *, char *);
void SelectAttachFile(HWND, char *, char *);
void AttachFileOpen(HWND, char *, char *);
BOOL CenterWindow(HWND, HWND);
int  MailPrint(HWND);
BOOL SmtpPop(SOCKET, char *, char *, char *);
int SaveMailData(LPSTR);
void Store_Mail(SOCKET, int, LPSTR);
void mid(char *, char *, int, int);
void SetLocalMail(HWND);
void Read_Mail(char *, HWND);
BOOL FileExists(LPSTR);
void LoadRegKey();
void SaveRegKey();
ATOM InitApp(HINSTANCE);
BOOL InitInstance(HINSTANCE, int);
BOOL InitFolder();
BOOL RcvMail(HWND, HWND);
void MessagePump();

// グローバル変数の定義
int  No, VersionNo;
int  sb_size;
char szWinName[] = "mailviewer";
char szUserName[256], szPass[256], szName[256], szMailAddr[256];
char szSmtpServer[256],  szPopServer[256], szMailData[MAIL_MAX];
char *body;
char szFileName[256], szFile[256] = "",	szDir[TEMP_MAX];
char szHeaderMsg[TEMP_MAX], szHeader[1024], strMsg[1024], strText[TEMP_MAX];
char szToData[TEMP_MAX], szSubjectData[TEMP_MAX], szBodyData[SEND_MAX], szFileData[TEMP_MAX];
char szAtFile[TEMP_MAX], szAtFileName[TEMP_MAX], *szAtFileData;
char RepSubject[1024], RepFrom[1024], RepDate[1024];
char MailSevePath[1024];
char szSelMailFileName[TEMP_MAX], szSelMailSubject[TEMP_MAX];

HINSTANCE hInst;
HWND      hParent;
BOOL      bRet, bApop, bReply, bPSmtp, bSMailDel;
SOCKET    s;
static HWND hStatus, hProgress;
DWORD dwCookie;

// ステータスバーに表示する説明
char szTipText[][256] = {
	"テキストファイルから添付ファイルを取り出します",
	"メールをテキスト形式で保存します",
	"印刷ページの設定をします(未実装)",
	"受信したメールを印刷します",
	"アプリケーションを終了します",
	"メールを作成して送信します",
	"メールを受信します",
	"選択中のメールに返信します",
	"選択中のメールを削除します",
	"ヘッダ情報を表示します",
	"メールサーバの設定を行います",
	"ヘルプを表示します",
	"バージョン情報を表示します",
};

typedef struct _tagSORTDATA{
	HWND hwndList;		// リストビューのHWND
	int isortSubItem;	// ソートするサブアイテムインデックス
	int iUPDOWN;		// 昇順か降順か
} SORTDATA;
