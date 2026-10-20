unit DmDatos;

interface

uses
  a3ERPActiveX_TLB, System.SysUtils, System.Classes, DLL_dmEstandar, Data.DB,
  Data.Win.ADODB, JncLib, Dll_Modelo, Dialogs, System.UiTypes;

type
  TdmdDatos = class(TdmDLL)
    qLineasLote: TADOQuery;
  private
    { Private declarations }
  public
    { Public declarations }
    function SumaCostesAdiciones(vIdAlbv:double):double;
    function SumaUnidadesLineas(vIdAlbv:double; ModificarCostesManuales:boolean):double;
    procedure ActualizaCosteLineas(vIdAlbv,vCosteUnidad:double; vModificarCostesManuales:boolean );
    procedure EliminaCostesAdicionalesAlbaran(vIdALbv:double);
    function BuscarParam(vIdalbv:double; vparametro:string):boolean;
    procedure ObtenerLineasConLote(vIdAlbc:double);
    function ObtenerCodPro(vIdAlbc:double):string;
    procedure ProcesarAlternativosAlbaran(const IdAlbC: Double);
    procedure GenerarRegularizacionesAlternativosVenta(const IdAlbv: Double);
   end;

var
  dmdDatos: TdmdDatos;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TdmdDatos }

procedure TdmdDatos.ActualizaCosteLineas(vIdAlbv, vCosteUnidad: double; vModificarCostesManuales:boolean);
var
  NaxAlbaran : albaran;
  NaxVarios : Varios;
  lUnidades : double;
  lCoste : double;
  lObtPrcCoste : string;
  lNUmLin : double;
  lCodArt, lCodAlm : string;
  lPrcCoste : double;
begin

  Naxvarios := coVarios.Create;
  NaxAlbaran := coAlbaran.Create;
  NaxAlbaran.Iniciar;
  NaxAlbaran.Modifica(vIdAlbv, False);

  with query do
  begin
    close;
    sql.Clear;

    sql.Add('Select codart, codalm, numlinalb, unidades, prcmedio,ObtPrcCoste from linealba with(nolock)');
    sql.Add('where idalbv = ' + formateafloatsql(vIdAlbv));

    open;

    first;

    while not eof do
    begin

      lNumlin :=fieldbyname('numlinalb').AsFloat;

      NaxAlbaran.EditarLinea(lNumLin);
      //lCoste := NaxAlbaran.AsFloatLin['PrcMedio'];
      lUnidades :=NaxAlbaran.AsFloatLin['Unidades'];
      lObtPrcCoste :=NaxAlbaran.AsStringLin['ObtPrcCoste'];
      lCodAlm := Fieldbyname('CodAlm').AsString;
      lCodArt := Fieldbyname('CodArt').AsString;
     

      if ((uppercase(lObtPrcCoste) <> 'MANU') or (vModificarCostesManuales)) then
      begin

        lPrcCoste := NaxAlbaran.AsFloatLin['SOL_COSTEANTERIOR'];

         //lPrcCoste := NaxVarios.ObtPrecioMedio(lCodArt,lCodAlm);
         if ((vCosteUnidad = 0) and  (lPrcCoste = 0) and (uppercase(lObtPrcCoste) <> 'MANU')) then
         begin

              lPrcCoste := NaxVarios.ObtPrecioCoste(lCodArt,lCodAlm,lObtPrcCoste);

              NaxAlbaran.AsFloatLin['PrcMedio'] := lPrcCoste +  vCosteUnidad;
              NaxAlbaran.AsStringLin['SOL_MODCOSTESADICIONALES'] := 'T';
         end
         else
         if ((vCosteUnidad <> 0) and  (lPrcCoste = 0) and (uppercase(lObtPrcCoste) <> 'MANU')) then
         begin

              lPrcCoste := NaxVarios.ObtPrecioCoste(lCodArt,lCodAlm,lObtPrcCoste);

              NaxAlbaran.AsFloatLin['PrcMedio'] := lPrcCoste +  vCosteUnidad;
              NaxAlbaran.AsStringLin['SOL_MODCOSTESADICIONALES'] := 'T';
         end
         else
         if uppercase(lObtPrcCoste) <> 'MANU' then
         begin

              NaxAlbaran.AsFloatLin['PrcMedio'] := lPrcCoste +  vCosteUnidad;
              NaxAlbaran.AsStringLin['SOL_MODCOSTESADICIONALES'] := 'T';
         end
         else
         if ((uppercase(lObtPrcCoste) = 'MANU')  and (lPrcCoste <> 0)) then
         begin

            NaxAlbaran.AsFloatLin['PrcMedio'] := lPrcCoste + vCosteUnidad ;
            NaxAlbaran.AsStringLin['SOL_MODCOSTESADICIONALES'] := 'T';
         end
         else
         begin

            lPrcCoste := NaxAlbaran.AsFloatLin['PrcMedio'];
           
            NaxAlbaran.AsFloatLin['PrcMedio'] := lPrcCoste + vCosteUnidad ;
            NaxAlbaran.AsStringLin['SOL_MODCOSTESADICIONALES'] := 'T';
         end;


      end;
      NaxAlbaran.AnadirLinea;

      next;
    end;

  end;
  aGuardadoPorDll[TTablaA3.tipoDocumentoPorParam('AV')] := 'T';
  NaxAlbaran.Anade;
  NaxAlbaran.Acabar;
end;

function TdmdDatos.BuscarParam(vIdalbv:double; vparametro: string): boolean;
begin
  with query do
  begin
    close;
    sql.Clear;

    sql.Add('Select coalesce(' + vparametro + ',''F'') Parametro from cabealbv with(nolock)');
    sql.Add('where idalbv = ' + formateafloatsql(vIdAlbv));

    open;

    result := FieldByName('parametro').asstring = 'T';
  end;
end;

procedure TdmdDatos.EliminaCostesAdicionalesAlbaran(vIdALbv: double);
begin
   with query do
  begin
    close;
    sql.Clear;

    sql.Add('Delete sol_costesadicionales where sol_idalbv = ' + formateafloatsql(vIdAlbv));

    execsql;
  end;
end;

procedure TdmdDatos.GenerarRegularizacionesAlternativosVenta(
  const IdAlbv: Double);
type
  TLinea = record
    CodAlm: string;
    CodArt: string;
    Unidades: Double; // POSITIVAS para la comparacion
  end;
var
  QDatos, QEstado, QExec: TADOQuery;
  Lineas: array of TLinea;
  I, J, N, Inicio, Fin: Integer;
  Almacen: string;
  TieneControl, Diferente, Reservado, Iniciado: Boolean;
  Estado: string;
  IdReg: Currency;
  Reg: IRegularizacion;
  Valores: Double;

  procedure SQLExec(const S: string);
  begin
    QExec.Close;
    QExec.SQL.Text := S;
    QExec.ExecSQL;
  end;

  procedure ParamControl(const AAlmacen: string);
  begin
    QEstado.Close;
    QEstado.SQL.Text :=
      'SELECT ESTADO,IDREG FROM dbo.SOL_REGALT_AV ' +
      'WHERE IDALBV=:ALB AND CODALM=:ALM';
    QEstado.Parameters.ParamByName('ALB').Value := IdAlbv;
    QEstado.Parameters.ParamByName('ALM').Value := AAlmacen;
    QEstado.Open;
  end;

begin
  if IdAlbv <= 0 then Exit;
  QDatos := TADOQuery.Create(nil);
  QEstado := TADOQuery.Create(nil);
  QExec := TADOQuery.Create(nil);
  try
    QDatos.Connection := query.Connection;
    QEstado.Connection := query.Connection;
    QExec.Connection := query.Connection;

    // Agrupar SOLO cantidades positivas del albaran ya guardado.
    // CODARTALT se conserva exactamente como se almacena (15 caracteres).
    QDatos.SQL.Text :=
      'SELECT L.CODALM, A.CODARTALT AS CODART, ' +
      'SUM(CAST(L.UNIDADES AS decimal(22,6)) * ' +
      'CAST(A.UNIPAQV AS decimal(22,6))) AS UNIDADES ' +
      'FROM LINEALBA L INNER JOIN ARTICULO A ON A.CODART=L.CODART ' +
      'WHERE L.IDALBV=:ALB AND L.UNIDADES > 0 ' +
      ' AND NULLIF(LTRIM(RTRIM(A.CODARTALT)), '''') IS NOT NULL ' +
      ' AND A.UNIPAQV > 0 ' +
      'GROUP BY L.CODALM,A.CODARTALT ' +
      'ORDER BY L.CODALM,A.CODARTALT';
    QDatos.Parameters.ParamByName('ALB').Value := IdAlbv;
    QDatos.Open;
    while not QDatos.Eof do
    begin
      N := Length(Lineas);
      SetLength(Lineas,N+1);
      Lineas[N].CodAlm := QDatos.FieldByName('CODALM').AsString;
      Lineas[N].CodArt := QDatos.FieldByName('CODART').AsString;
      Lineas[N].Unidades := QDatos.FieldByName('UNIDADES').AsFloat;
      QDatos.Next;
    end;
    QDatos.Close;

    Inicio := 0;
    while Inicio < Length(Lineas) do
    begin
      Almacen := Lineas[Inicio].CodAlm;
      Fin := Inicio;
      while (Fin < Length(Lineas)) and
        (Lineas[Fin].CodAlm = Almacen) do Inc(Fin);

      if Trim(Almacen) = '' then
        raise Exception.Create('Almacen vacio en AV ' + FloatToStr(IdAlbv));

      ParamControl(Almacen);
      TieneControl := not QEstado.Eof;
      if TieneControl then Estado := QEstado.FieldByName('ESTADO').AsString
      else Estado := '';
      QEstado.Close;

      if TieneControl then
      begin
        if Estado <> 'C' then
          raise Exception.Create('Regularizacion de AV ' + FloatToStr(IdAlbv) +
            ', almacen ' + Trim(Almacen) +
            ' pendiente de comprobacion. NO reintentar automaticamente.');

        // Si cambia el documento despues de regularizar, bloquear nuevos
        // descuentos y exigir ajuste manual controlado.
        Diferente := False;
        QEstado.SQL.Text :=
          'SELECT CODART,UNIDADES FROM dbo.SOL_REGALT_DET ' +
          'WHERE IDALBV=:ALB AND CODALM=:ALM ORDER BY CODART';
        QEstado.Parameters.ParamByName('ALB').Value := IdAlbv;
        QEstado.Parameters.ParamByName('ALM').Value := Almacen;
        QEstado.Open;
        J := Inicio;
        while (not QEstado.Eof) and (J < Fin) do
        begin
          Valores := QEstado.FieldByName('UNIDADES').AsFloat;
          if (QEstado.FieldByName('CODART').AsString <> Lineas[J].CodArt) or
             (Abs(Valores - Lineas[J].Unidades) > 0.00001) then
            Diferente := True;
          Inc(J);
          QEstado.Next;
        end;
        if (not QEstado.Eof) or (J <> Fin) then Diferente := True;
        QEstado.Close;
        if Diferente then
          raise Exception.Create('AV ' + FloatToStr(IdAlbv) +
            ' modificado tras la regularizacion del almacen ' + Trim(Almacen) +
            '. Revisar y realizar ajuste diferencial MANUAL.');
      end
      else
      begin
        // Reservar antes de crear el documento ActiveX. Esta reserva
        // es persistente y bloquea un reintento despues de un fallo.
        QExec.SQL.Text :=
          'INSERT INTO dbo.SOL_REGALT_AV (IDALBV,CODALM,ESTADO,FECHA) ' +
          'VALUES (:ALB,:ALM,''P'', :FECHA)';
        QExec.Parameters.ParamByName('ALB').Value := IdAlbv;
        QExec.Parameters.ParamByName('ALM').Value := Almacen;
        QExec.Parameters.ParamByName('FECHA').Value := Now;
        QExec.ExecSQL; // PK impide dos procesos concurrentes
        Reservado := True;
        for I := Inicio to Fin - 1 do
        begin
          QExec.SQL.Text :=
            'INSERT INTO dbo.SOL_REGALT_DET ' +
            '(IDALBV,CODALM,CODART,UNIDADES) ' +
            'VALUES (:ALB,:ALM,:ART,:UNI)';
          QExec.Parameters.ParamByName('ALB').Value := IdAlbv;
          QExec.Parameters.ParamByName('ALM').Value := Almacen;
          QExec.Parameters.ParamByName('ART').Value := Lineas[I].CodArt;
          QExec.Parameters.ParamByName('UNI').Value := Lineas[I].Unidades;
          QExec.ExecSQL;
        end;

        Reg := nil;
        Iniciado := False;
        try
          Reg := CoRegularizacion.Create;
          Reg.Iniciar;
          Iniciado := True;
          Reg.Nuevo(FormatDateTime('dd/mm/yyyy', Date), Almacen);
          for I := Inicio to Fin - 1 do
          begin
            Reg.NuevaLineaArt(Lineas[I].CodArt, -Lineas[I].Unidades);
            Reg.AnadirLinea;
          end;
          IdReg := Reg.Anade;
          if IdReg <= 0 then
            raise Exception.Create('Regularizacion sin identificador valido');

          // No borrar nunca una reserva P por un fallo.
          QExec.SQL.Text :=
            'UPDATE dbo.SOL_REGALT_AV SET ESTADO=''C'',IDREG=:REG ' +
            'WHERE IDALBV=:ALB AND CODALM=:ALM AND ESTADO=''P''';
          QExec.Parameters.ParamByName('REG').Value := IdReg;
          QExec.Parameters.ParamByName('ALB').Value := IdAlbv;
          QExec.Parameters.ParamByName('ALM').Value := Almacen;
          QExec.ExecSQL;
        except
          on E: Exception do
          begin
            if Iniciado then
              try Reg.Cancela; except end;
            // El estado P permanece: puede haber documento ya generado.
            raise Exception.Create('AV ' + FloatToStr(IdAlbv) +
              ', almacen ' + Trim(Almacen) + ': ' + E.Message +
              '. Estado P: verificar antes de reintentar.');
          end;
        end;
        if Iniciado then
          try Reg.Acabar; except end;
        Reg := nil;
      end;
      Inicio := Fin;
    end;
  finally
    QExec.Free;
    QEstado.Free;
    QDatos.Free;
  end;


end;

function TdmdDatos.ObtenerCodPro(vIdAlbc: double): string;
begin
   with query do
   begin
     close;
     sql.Clear;

     sql.Add('Select p.Param1 from cabealbc c with(nolock)');
     sql.Add('left outer join proveed p with(nolock) on p.codpro = c.codpro');
     sql.add(' where c.idalbc = ' + formateafloatsql(vIdAlbc));

     open;

     result := fieldbyname('Param1').AsString;
   end;
end;

procedure TdmdDatos.ObtenerLineasConLote(vIdAlbc: double);
var
  lcadena:string;
begin
   with qLineasLote do
   begin
     close;
     sql.Clear;

     // Obtenemos todas las líneas con lotes, que habría que actualizar.
     sql.Add('Select l.idlin from linealba as l with(nolock) ');
     sql.Add(' left outer join articulo with(nolock) on articulo.codart = l.codart ');
     sql.Add(' where articulo.haylotes = ''T'' ');
     sql.Add(' and ltrim(l.lote) = '''' ');
     sql.Add(' and l.idalbc= ' + formateafloatsql(vIdAlbc)+ ' and l.unidades > 0 ');


     lCadena := sql.Text;
     open;

   end;
end;

function TdmdDatos.SumaCostesAdiciones(vIdAlbv: double): double;
begin
  with query do
  begin
    close;
    sql.Clear;

    sql.Add('Select sum(coalesce(sol_coste,0)) coste from sol_costesadicionales with(nolock)');
    sql.Add('where sol_idalbv = ' + formateafloatsql(vIdAlbv));
    sql.Add('and coalesce(sol_sumar,''F'') = ''T''');

    open;

    result := fieldbyname('coste').AsFloat;

  end;
end;

function TdmdDatos.SumaUnidadesLineas(vIdAlbv: double; ModificarCostesManuales:boolean): double;
begin
    with query do
  begin
    close;
    sql.Clear;

    sql.Add('Select sum(coalesce(unidades,0)) unidades from linealba with(nolock)');
    sql.Add('where idalbv = ' + formateafloatsql(vIdAlbv));
    if not ModificarCostesManuales then
         sql.add('and obtprccoste <> ''MANU''');

    open;

    result := fieldbyname('unidades').AsFloat;

  end;
end;


// NUEVO: trabaja sobre un AV ya confirmado. Nunca ejecuta UPDATE de documentos.
procedure TdmdDatos.ProcesarAlternativosAlbaran(const IdAlbC: Double);
type
  TGrupo = record
    IdPedC: Double;
    CodAlternativo: string;
    Necesarias: Double;
    YaEnAlbaran: Double;
  end;
  TLineaPlan = record
    NumLinPed: Double;
    Cantidad: Double;
  end;
var
  QGrupos, QLineas: TADOQuery;
  Grupos: array of TGrupo;
  Plan: array of TLineaPlan;
  NaxAlbaran: Albaran;
  I, J, N: Integer;
  Necesarias, Disponibles, Restantes, Pend, Cantidad: Double;
  Preparado, MarcaGuardado: Boolean;
  Aviso: string;
  function Num(const Valor: Double): string;
  begin
    Result := FormatFloat('0.######', Valor);
  end;
begin
  if IdAlbC <= 0 then Exit;
  QGrupos := TADOQuery.Create(nil);
  QLineas := TADOQuery.Create(nil);
  try
    QGrupos.Connection := query.Connection;
    QLineas.Connection := query.Connection;
    QGrupos.SQL.Text :=
      'SELECT L.IDPEDC, A.CODARTALT, ' +
      ' SUM(L.UNIDADES * A.UNIPAQC) AS NECESARIAS, ' +
      ' (SELECT COALESCE(SUM(X.UNIDADES),0) FROM LINEALBA X ' +
      '   WHERE X.IDALBC = :ALB2 AND X.IDPEDC = L.IDPEDC ' +
      '     AND X.CODART = A.CODARTALT) AS YA_EN_ALBARAN ' +
      'FROM LINEALBA L INNER JOIN ARTICULO A ON A.CODART = L.CODART ' +
      'WHERE L.IDALBC = :ALB AND L.IDPEDC IS NOT NULL ' +
      ' AND NULLIF(LTRIM(RTRIM(A.CODARTALT)), '''') IS NOT NULL ' +
      ' AND L.UNIDADES > 0 ' +
      'GROUP BY L.IDPEDC, A.CODARTALT';
    QGrupos.Parameters.ParamByName('ALB').Value := IdAlbC;
    QGrupos.Parameters.ParamByName('ALB2').Value := IdAlbC;
    QGrupos.Open;
    while not QGrupos.Eof do
    begin
      N := Length(Grupos);
      SetLength(Grupos, N+1);
      Grupos[N].IdPedC := QGrupos.FieldByName('IDPEDC').AsFloat;
      Grupos[N].CodAlternativo := QGrupos.FieldByName('CODARTALT').AsString;
      Grupos[N].Necesarias := QGrupos.FieldByName('NECESARIAS').AsFloat;
      Grupos[N].YaEnAlbaran := QGrupos.FieldByName('YA_EN_ALBARAN').AsFloat;
      QGrupos.Next;
    end;
    QGrupos.Close;

    for I := 0 to High(Grupos) do
    begin
      Necesarias := Grupos[I].Necesarias - Grupos[I].YaEnAlbaran;
      if Necesarias <= 0.000001 then Continue;
      if Grupos[I].Necesarias <= 0 then Continue;

      QLineas.Close;
      QLineas.SQL.Text :=
        'SELECT NUMLINPED, ' +
        ' (UNIDADES-COALESCE(UNISERVIDA,0)-COALESCE(UNIANULADA,0)) AS PENDIENTES ' +
        'FROM LINEPEDI WHERE IDPEDC = :PED AND CODART = :ART ' +
        ' AND UNIDADES-COALESCE(UNISERVIDA,0)-COALESCE(UNIANULADA,0) > 0 ' +
        'ORDER BY NUMLINPED';
      QLineas.Parameters.ParamByName('PED').Value := Grupos[I].IdPedC;
      QLineas.Parameters.ParamByName('ART').Value := Grupos[I].CodAlternativo;
      QLineas.Open;
      SetLength(Plan,0);
      Restantes := Necesarias;
      Disponibles := 0;
      while not QLineas.Eof do
      begin
        Pend := QLineas.FieldByName('PENDIENTES').AsFloat;
        Disponibles := Disponibles + Pend;
        if (Restantes > 0.000001) and (Pend > 0) then
        begin
          if Pend < Restantes then Cantidad := Pend else Cantidad := Restantes;
          N := Length(Plan);
          SetLength(Plan,N+1);
          Plan[N].NumLinPed := QLineas.FieldByName('NUMLINPED').AsFloat;
          Plan[N].Cantidad := Cantidad;
          Restantes := Restantes - Cantidad;
        end;
        QLineas.Next;
      end;
      QLineas.Close;
      if Restantes > 0.000001 then
      begin
        Aviso := 'No se puede servir el articulo alternativo ' +
          Trim(Grupos[I].CodAlternativo) + ' del pedido ' + Num(Grupos[I].IdPedC) + '.' + sLineBreak +
          'Necesarias: ' + Num(Necesarias) + '. Pendientes: ' + Num(Disponibles) + '.' + sLineBreak +
          'El albaran principal se conserva y NO se sirve ninguna linea del alternativo.';
        MessageDlg(Aviso, mtWarning, [mbOK], 0);
        Continue;
      end;

      NaxAlbaran := nil;
      Preparado := False;
      MarcaGuardado := False;
      try
        NaxAlbaran := CoAlbaran.Create;
        NaxAlbaran.Iniciar;
        Preparado := True;
        NaxAlbaran.Modifica(IdAlbC, True);
        NaxAlbaran.IniciarServir('P', Grupos[I].IdPedC, False);
        for J := 0 to High(Plan) do
          NaxAlbaran.ServirLinea(0, Plan[J].NumLinPed, 0, 0,
            Plan[J].Cantidad, 0, '', '', '', '');
        NaxAlbaran.FinServir;
        // MISMO patron de recursividad que ActualizaCosteLineas.
        aGuardadoPorDll[TTablaA3.tipoDocumentoPorParam('AC')] := 'T';
        MarcaGuardado := True;
        NaxAlbaran.Anade;
      except
        on E: Exception do
        begin
          if MarcaGuardado then
            aGuardadoPorDll[TTablaA3.tipoDocumentoPorParam('AC')] := 'F';
          if Assigned(NaxAlbaran) and Preparado then
            try NaxAlbaran.Cancela; except end;
          MessageDlg('No se ha podido completar el servido del alternativo ' +
            Trim(Grupos[I].CodAlternativo) + ' en el albaran ' + Num(IdAlbC) +
            '.' + sLineBreak + E.Message + sLineBreak +
            'Verifique las lineas antes de reintentar.', mtError, [mbOK], 0);
        end;
      end;
      if Assigned(NaxAlbaran) and Preparado then
        try NaxAlbaran.Acabar; except end;
    end;
  finally
    QLineas.Free;
    QGrupos.Free;
  end;
end;

end.
