///////////////////////////////////////////////////////////////////////////////////////
// Mail Viewer Version 0.07
// �쐬��:���q����(Angelic Software)
// �X�V�o�[�W����:0.07(2007�N12��27��)
// http://www.�V�g�̂�����.com/
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

#define UP          1			// ����
#define DOWN        2			// �~��
#define NO_OF_SUBITEM 5			// �T�u�A�C�e���̐�

// �E�B���h�E�v���V�[�W���̌^�錾
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK NewMailProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MailSetProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK HeaderProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK AboutProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK AttachFileProc(HWND, UINT, WPARAM, LPARAM);

// �t�H���_�I���E�B���h�E�v���V�[�W���̌^�錾
int CALLBACK SHMyProc(HWND, UINT, LPARAM, LPARAM);

// ���X�g��r�p�̃v���V�[�W���̌^�錾
int CALLBACK ListCompProc(LPARAM, LPARAM, LPARAM);

// �֐��̃v���g�^�C�v�錾
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

// �O���[�o���ϐ��̒�`
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

// �X�e�[�^�X�o�[�ɕ\���������
char szTipText[][256] = {
	"�e�L�X�g�t�@�C������Y�t�t�@�C�������o���܂�",
	"���[�����e�L�X�g�`���ŕۑ����܂�",
	"����y�[�W�̐ݒ�����܂�(������)",
	"��M�������[����������܂�",
	"�A�v���P�[�V�������I�����܂�",
	"���[�����쐬���đ��M���܂�",
	"���[������M���܂�",
	"�I�𒆂̃��[���ɕԐM���܂�",
	"�I�𒆂̃��[�����폜���܂�",
	"�w�b�_����\�����܂�",
	"���[���T�[�o�̐ݒ���s���܂�",
	"�w���v��\�����܂�",
	"�o�[�W��������\�����܂�",
};

typedef struct _tagSORTDATA{
	HWND hwndList;		// ���X�g�r���[��HWND
	int isortSubItem;	// �\�[�g����T�u�A�C�e���C���f�b�N�X
	int iUPDOWN;		// �������~����
} SORTDATA;
