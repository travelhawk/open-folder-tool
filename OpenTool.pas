unit Projekt1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ShlObj, ActiveX, ShellApi, StdCtrls, XPMan, ComCtrls;
  //ShlObj, ActiveX, ShellApi müssen eingebunden werden

type
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Label1: TLabel;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    Button14: TButton;
    XPManifest1: TXPManifest;
    StatusBar1: TStatusBar;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure ShowSpecialFolder(const AFolder: Integer); //Ordner mittels Windowskonstante anzeigen
var
  ItemIDList: PItemIDList;
  ShExInfo: ShellExecuteInfo;
  ShellH: IMalloc;
begin
  if SUCCEEDED(ShGetSpecialFolderLocation(Application.Handle, AFolder, ItemIDList)) then
  begin
    try
      FillChar(ShExInfo, SizeOf(ShExInfo), 0);
      with ShExInfo do
      begin
        cbSize := SizeOf(ShExInfo);
        nShow := SW_SHOW;
        fMask := SEE_MASK_IDLIST;
        lpIDList := ItemIDList;
      end;
      ShellExecuteEx(@ShExInfo);
    finally
      if SHGetMalloc(ShellH) = NOERROR then
        ShellH.Free(ItemIDList);
    end;
  end
  else
    RaiseLastOSError;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_Desktop);
//weitere Konstanten sind in der ShlObj.pas zu finden
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_Drives);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_Controls);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_Network);
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_Printers);
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_STARTUP);
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_RECENT);
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_SENDTO);
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_BITBUCKET);
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_STARTMENU);
end;

procedure TForm1.Button11Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_FONTS);
end;

procedure TForm1.Button12Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_TEMPLATES);
end;

procedure TForm1.Button13Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_COOKIES);
end;

procedure TForm1.Button14Click(Sender: TObject);
begin
ShowSpecialFolder(CSIDL_HISTORY);
end;

end.
 