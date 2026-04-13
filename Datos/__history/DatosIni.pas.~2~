unit DatosIni;

interface

uses dll_fichIni, System.SysUtils, Dialogs, System.Classes;

type
   //Clase encargada de leer los datos de un fichero ini.
  TDatosIniAplicacion = class(TDllDatosIni)
  private
    FModificarCostesManuales: boolean;


   public
    procedure LeeIni; override;


    property ModificarCostesManuales : boolean read  FModificarCostesManuales write FModificarCostesManuales;




  end;

implementation

{ TDatosIniAplicacion }

procedure TDatosIniAplicacion.LeeIni;
begin
  inherited;


  ModificarCostesManuales:= FArchivoIni.ReadString('Configuracion', 'ModificarCostesManuales', 'F') = 'T';


end;

end.
