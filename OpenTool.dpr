program Project1;

uses
  Forms,
  Projekt1 in 'Projekt1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
