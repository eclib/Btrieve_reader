'----------------------------------------------------------------------------
'PROGRAMA: SuperEXPorta def v2
'DESCRIPCION: Exporta a un archivo .csv cualquier archivo btrieve contenido en trans.dat
'AUTORES: Manuel Sotomayor - Eduardo Llanquileo
'FECHA: Marzo 2007
'--------------------------------------------------------------------------------

'    ---------------------------------------------------------
DIM  shared buf$, posblock$      'variables que requieren asignacion estatica de memoria
DIM  shared file.btrieve
dim  shared transfea,transfec,transloc,local$
DIM  shared directorio.datos$,archivo.salida$
DIM  shared Btipo$(15)           'Tipo de datos
DIM  shared Field.id(900),Field.File(900),Field.name$(900),Field.type(900),Field.offset(900),Field.size(900),Field.dec(900)
DIM  shared TField.id(100),tField.File(100),tField.name$(100),tField.type(100),tField.offset(100), tField.size(100),tField.dec(100)
DIM shared vfiles$(1000) 'para leer los archivos

DIM  shared bfiles(200) ' para llevar la cuenta de cuantas veces se ha exportado el mismo archivo
' Definiciones de tipos de datos
Btipo$(0)="String":Btipo$(1)="Integer":Btipo$(2)="Float":Btipo$(3)="Date":Btipo$(4)="Time":Btipo$(5)="Decimal"
Btipo$(6)="Money":Btipo$(7)="Logical":Btipo$(8)="Numeric":Btipo$(9)="Bfloat":Btipo$(10)="Lstring":Btipo$(11)="Zstring"
Btipo$(12)="Note":Btipo$(13)="Lvar"::Btipo$(14)="Unsign Binary":Btipo$(15)="Autoincremen"
'------ Definiciones nemotecnicas de operaciones mas utilizadas ----------------
BtrOpen%=0         'Abre archivo
BtrGetFirst%=12    'Se posiciona en el primer registro, orden llave
BtrGetNext%=6      'Lee siguiente registro, orden llava
BtrFirst%=33       'Se posiciona en el primer registro, orden fisico
BtrNext%=24        'Lee siguiente registro, orden fisico
BtrCloseAll%=28    'Cierra archivos
'--------------------------------------------------------------------------------
'    ---------------------------------------------------------
for i=1 to 200: bfiles(i)=0: next i
' -------------------------------------------------------------------------------
' ------------------- FUNCIONES DE APOYO ------------------------------------------
'
  
FUNCTION xdir (idir$) 
   ' ---- busca los archivos y los deja en un vector
   ' ---- Admite un maximo de 500 archivos
   '  --- devuelve la cantidad de archivos encontrados
   spec$="dir /b /on " + idir$ + " > temp.tmp"
   'print spec$
   shell spec$
   open "temp.tmp" for input as #1
   i=0
   while not eof(1)
       i=i+1
       input #1, linea$
	 'if i > 100 then exit do
	  vfiles$(i)=linea$
	  
   wend
   close #1
   kill "temp.tmp"
   xdir=i
END FUNCTION

FUNCTION EXTRAE.FILE$(xinput$)
   ' extrae el nombre del archivo  a partir de la entrada proporcionada
   ' devuelve parametro FATAL si no existe parametro 
   ' o existe parametro sin datos
   
   xtmp$=ucase$(xinput$)
   p=instr(xtmp$, "/FILE=")
   if p=0 then extrae.file$="%FATAL%":exit function
   file$=""
   for i=p+6 to len(xtmp$)
      if mid$(xtmp$,i,1)<>"/" then file$=file$+mid$(xtmp$,i,1)
      if mid$(xtmp$,i,1)="/"  then exit for	  
   next i
   if file$="" then file$="%FATAL%"
   extrae.file$=file$
END FUNCTION

FUNCTION EXTRAE.DIR$(xinput$)
   ' extrae el nombre del directorio donde buscar a partir de la entrada proporcionada
  
   xtmp$=ucase$(xinput$)
   p=instr(xtmp$, "/DIR=")
   if p=0 then extrae.dir$="":exit function ' si no hay parametro, se asume directorio acutal
   dir$=""
   for i=p+5 to len(xtmp$)
      if mid$(xtmp$,i,1)<>"/" then dir$=dir$+mid$(xtmp$,i,1)	  
	  if mid$(xtmp$,i,1)="/"  then exit for	
   next i
   
   if dir$="" then extrae.dir$="":directorio.datos$="":exit function
   if right$(dir$,1)<>"\" then dir$=dir$+"\"
   directorio.datos$=dir$
   extrae.dir$=dir$
   
END FUNCTION

FUNCTION EXTRAE.LOCAL$(xinput$)
   ' extrae el numero del  local a partir de la entrada m
  
   xtmp$=ucase$(xinput$)
   p=instr(xtmp$, "/LOCAL=")
   if p=0 then extrae.local$="":exit function ' si no hay parametro, se asume directorio acutal
   loc$=""
   for i=p+7 to len(xtmp$)
      if mid$(xtmp$,i,1)<>"/" then loc$=loc$+mid$(xtmp$,i,1)	  
	  if mid$(xtmp$,i,1)="/"  then exit for	
   next i
   'print "local:", loc$
   if loc$="" then extrae.local$="000":exit function
  
   extrae.local$=ltrim$(rtrim$(loc$))
   
END FUNCTION

FUNCTION MUESTRA.HELP
  print "SEXP /FILE=archivo /DIR=directorio /LOCAL=local /LOG= archivo"
  print "       Archivo: cualquier especificacion validad,incluye comodines"
  
END FUNCTION

function  PARSE.ENTRADA(entrada$)
     if rtrim$(entrada$)="" then x=muestra.help:end ' No se admiten parametros nulos
	 dir$=extrae.dir$(entrada$)	  'Determina  la especificacion de directorio, si existe
	 file$=extrae.file$(entrada$)	 
     if file$="%FATAL%" then print "### Error: indicar archivo" : end	 
	 x=xdir(dir$+file$) ' extrae los archivo especificados
	 if x =0 then print "### Error: no se encontro archivo ":end
	 parse.entrada=x 'devuelve la cantidad de archivos leidos
	 local$=extrae.local$(entrada$)
end function           

FUNCTION limpia$(arc$)
	sal$=""
	
	for i = 1 to len(arc$)
	   x$= mid$(arc$,i,1)
	   if x$="'" or x$="," or x$=CHR$(34) then x$= " "
	   sal$=sal$+x$
	next i   
	limpia$=sal$
END FUNCTION

FUNCTION VERIFICA.BTRIEVE$(archivo$,file.btrieve)
'    --------------------------------------------------------- 
       open "trans.dat" for input as #6
       i=0
	   'print "...."+archiv$
	   ix%=0
	   xarchiv$=archivo$
	   
        while not eof(6)
		'i=i+1
            input #6,transId%   'ID del archivo trans.dat
            input #6,transRe$   'Expresion a buscar en trans.dat
            input #6,transBu$   'Archivo del cual debo sacar su estructura
			input #6,transFEA2  'Defino si va hay que añadir la columna MOVFECHA y la saco del nombre de archivo
			input #6,transFEC2  'Defino si va hay que añadir la columna MOVFECHA y la creo
			input #6,transLOC2  'defino si hay que añadir la columna MOVLOCAL
			     
            if INSTR(xarchiv$,rtrim$(transRe$)) <> 0 then
                ix% = transid%
                archsal$ = rtrim$(transBu$)
				transFEA=transFEA2
				transFEC=transFEC2
				transLOC=transLOC2
				'print "###",xarchiv$,archiv$,transbu$
			end if
        WEND
		'print "salida archivo:",archivo$,archsal$,ix%
		if ix%=0 then archsal$="### - Archivo no btrieve "
		file.btrieve=ix% ' asigna a variable global
		'print "file btrieve:", file.btrieve
		verifica.btrieve$=archsal$
        close #6
END FUNCTION

'=======================================================================================================================================
' ============================================AQUI  PARTE LA EJECUCION DEL PROGRAMA ================================================
'2) Lee diccionario de campos

cls
GOSUB TRASPASA.FIELD


entrada$=command$

  x=parse.entrada(entrada$)
  for xi=1 to x
    archivo.salida$=verifica.btrieve$(vfiles$(xi),file.btrieve)
 
	print xi;" Exporta:";vfiles$(xi);" DDF:";archivo.salida$;
	if file.btrieve<>0 then
	   archivo.original$=vfiles$(xi)  'trae el nombre original del archivo btrieve a traspasar
 	   GOSUB TRASPASA.TEMPORAL
	   GOSUB TRASPASA.FILES
	end if

  next xi
END



'poner el nombre del archivo original

TRASPASA.TEMPORAL:
'3) Traspasa los datos del id (ix) a una tabla temporal
largo.reg%=0
fd%=0
for i%=1 to maxfield%
    if field.file(i%)=file.btrieve then
        fd%=fd%+1
        tfield.id(fd%)=field.id(i%)
        tfield.name$(fd%)=field.name$(i%)   
        tfield.size(fd%)=field.size(i%)
        tfield.offset(fd%)=field.offset(i%)
        tfield.type(fd%)= field.type(i%)
        largo.reg%=largo.reg%+field.size(i%)
        tfield.dec(fd%)=field.dec(i%)
       ' print "fd:",fd%,tfield.name$(fd%),tfield.type(fd%),tfield.offset(fd%)
    end if
next i%
'print "LARGO:",largo.reg%,archivo.original$

lx%=largo.reg%
RETURN

TRASPASA.FILES:
   archx$=directorio.datos$+archivo.original$
   
   'print "archivo a leer:",archx$
   'print "archiv  salida:", archivo.salida$
   archivo.salida$=mid$(archivo.salida$,1,len(archivo.salida$)-3)+"CSV"
   print " CSV:"; archivo.salida$;
   
   GOSUB ABRE.ARCHIVOS
   GOSUB TRASPASA.ARCHIVOS
   GOSUB CIERRA.ARCHIVOS
RETURN

END


TRASPASA.ARCHIVOS:
'    ---------------------------------------------------------
        if bfiles(file.btrieve)=0 then open archivo.salida$ for output as #2  'determina si es la primera vez que abre el archivo
		if bfiles(file.btrieve)>0 then open archivo.salida$ for append as #2 
		
		
        'open archivo.salida$ for append as #2
		
        Leidos=0
        grabados=0
        op% = BtrFirst%: llave$ = STRING$(2, 1): regnum% = 0: GOSUB BTRV
        i%=0
		if bfiles(file.btrieve)=0 then 
		
		   if transLOC = 1 then print #2,"'MOVLOCAL ',";
		   if transFEA = 1 or transFEC = 1 then print #2,"'MOVFECHA ',";
		
           for i=1 to fd%
               print #2, "'"+tfield.name$(i),"',";
		   'print i, tfield.name$(i)
           next i
            print #2,""
	    end if
		
		'print "aca 5"	
        WHILE st% = 0
              leidos=leidos+1         
              i%=i%+1
			  
			  if transLOC = 1 then
					xsal$="'"+local$+"'"
					print #2,xsal$;",";
			  end if
			  
			  if transFEA = 1 then
					xsal$="'20"+mid$(archivo.original$,1,6)+"'"
					print #2,xsal$;",";
			  end if
				
			  if transFEC = 1 then
					xsal$="'071905'"
					print #2,xsal$;",";
			  end if
			  
			  
			  
              for f%=1 to fd%
                  fsize=tfield.size(f%)
                  ftype=tfield.type(f%)
                  field$=mid$(buf$,tfield.offset(f%)+1,tfield.size(f%))
                     select case ftype
                      case 0
                          xsal$=field$
                          if asc(mid$(xsal$,1,1))< 32 or asc(mid$(xsal$,1,1)) > 128 then xsal$=""
                          xsal$="'"+limpia$(xsal$)+"'"
                      case 1
                           if fsize=1 then xsal$=str$(asc(field$))
                           if fsize=2 then xsal$=str$(cvi(field$))
                           if fsize=4 then xsal$=str$(cvl(field$))
                      case 2
                           if fsize=4 then xsal$=str$(cvs(field$))
                           if fsize=8 then xsal$=str$(cvd(field$))
                     case 3
                         dia=asc(mid$(field$,1,1))
                         mes=asc(mid$(field$,2,1))
                         dia$=ltrim$(str$(dia)):mes$=ltrim$(str$(mes))
                             if dia < 10 then dia$="0"+ltrim$(dia$)
                             if mes < 10 then mes$="0"+ltrim$(mes$)
                         year$=ltrim$(str$(cvi(mid$(field$,3,2))))
                         xsal$="'"+year$+mes$+dia$+"'"
                      case else
                           'xsal$="'**"+field$+str$( tfield.type(fd%))+" "+STR$(FD%)+"'"
						   xsal$="' '"
                 end select          
                 print #2,xsal$;",";
              next f%     
              print #2,""
              grabados=grabados+1
              'LOCATE 12, 15: PRINT USING "& ###,###"; "Leidos       : "; Leidos
              'LOCATE 13, 15: PRINT USING "& ###,###"; "Grabados     : "; grabados
			  
              op% = BtrNext%: Gosub btrv
			  
        WEND
		print " leidos:"; leidos
  op% = BtrCloseAll%: GOSUB BTRV
  close #2
  bfiles(file.btrieve)=bfiles(file.btrieve)+1
RETURN



TRASPASA.FIELD:
'    ---------------------------------------------------------

		open "field.dat" for input as #6
		
        i=0
        while not eof(6)
			i%=i%+1
            input #6,xfield.id
            input #6,xfield.file
            input #6,xfield.name$
			input #6,xfield.type
			input #6,xfield.offset
			input #6,xfield.size
			input #6,xfield.dec  
			
			Field.id(i%)=xfield.id
            Field.file(i%)=xfield.file
            Field.name$(i%)=xfield.name$
            Field.type(i%)=xfield.type
            Field.offset(i%)=xfield.offset
            Field.size(i%)=xfield.size
            Field.dec(i%)=xfield.dec  

        WEND
		close #6
		maxfield%=i% 'maxima cantidad de datos leidos

RETURN

'===============================================================================================
ABRE.ARCHIVOS:
      
        buf$=space$(lx%)
        op% = BtrOpen%: posblock$ = SPACE$(128): llave$ = archx$+" "
        Gosub Btrv
     
        IF st% <> 0 THEN CALL btrv.err(op%, st%, llave$, 1, regnum%):print "error btrieve"+str$(St%):STOP

RETURN

'===============================================================================================
CIERRA.ARCHIVOS:
  op% = BtrCloseAll%: GOSUB BTRV
RETURN

'===============================================================================================
BTRV:
' Ejecuta las llamadas al motor btrieve y recibe los parametros segun las operaciones necesarias
 fcb% = SADD(buf$)- 188   'No se que hace
 CALL BTRV(op%, st%, posblock$, fcb%, lx%, llave$, regnum%)
RETURN