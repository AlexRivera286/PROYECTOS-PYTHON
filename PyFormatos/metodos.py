#METODOS

def date():
  date=datetime.now()
  dia=date.day
  mes=date.month
  anio=date.year
  if(mes==1):
    mes="enero"
  if(mes==2):
    mes="febrero"
  if(mes==3):
    mes="marzo"
  if(mes==4):
    mes="abril"
  if(mes==5):
    mes="mayo"
  if(mes==6):
    mes="junio"
  if(mes==7):
    mes="julio"
  if(mes==8):
    mes="agosto"
  if(mes==9):
    mes="septiembre"
  if(mes==10):
    mes="octubre"
  if(mes==11):
    mes="noviembre"
  if(mes==12):
    mes="diciembre"
  return f"{dia} de {mes} del {anio}"

#DEFINICION DEL METODO PARA LA CREACION DE LA ACTA
def acta(mes,acta):     #acta(mes,"acta-1")
  global cantidad
  nuevo=[]
  ticket=reconocimiento_acta(mes,acta,"ticket")
  datos=reconocimiento_acta(mes,acta,"datos")
  if(len(ticket)==1):
    if(re.search("-",ticket[0])):
      tickets=re.split("-",ticket[0])
      nuevo.append(tickets[0])
      nuevo.append(tickets[1])
  if(len(datos)==3):
    for x in range(len(datos)):
      nuevo.append(datos[x][0])
  if(len(nuevo)==5):
    #def nuevo_renglon(folio,usuario,ubicacion,serie):
    nuevo_renglon((nuevo[0]+"-"+nuevo[1]),nuevo[2],nuevo[3],nuevo[4])
    #Llenar arreglo datos
    datos_acta[cantidad]={
        'FGR':nuevo[0],
        'THEOS':nuevo[1],
        'USUARIO':nuevo[2],
        'CEDE':nuevo[3],
        'SERIE':nuevo[4],
        'GUIA':num_guia
    }
  cantidad+=1
  #print(cantidad)
  #Agregar excel Acta
  Excel_Actas(datos_acta)

  #guardar cambios al docx
  documento.save(f"drive/MyDrive/PyFormatos/Nuevo/ACUSE BC,{mes}2026.docx")

  return datos_acta

def reconocimiento_acta(mes,acta,tipo): # ("abril","acta-1","ticket")
  arreglo=[]
  v_nombre=[]
  v_cede=[]
  v_serie=[]
  v_theos=""
  v_fgr=""
  #Lee imagen del acta
  doc = fitz.open(f"drive/MyDrive/PyFormatos/Archivos/{mes}2026/{acta}.pdf") # Lectura del PDF
  modelo=ocr_predictor(det_arch='db_resnet50',reco_arch='crnn_vgg16_bn',pretrained=True,assume_straight_pages=True)
  page = doc[0] # Seleccion de la hoja del archivo PDF
  zoom = 2  # Renderizar como imagen (controlas resolución aquí)
  mat = fitz.Matrix(zoom, zoom)
  pix = page.get_pixmap(matrix=mat,dpi=600)
  # Convertir a una imagen modelo RGB (3 Colores)
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
  # utiliza el metodo crop para delimitar la imagen (izquierda, arriba, derecha, abajo).
  img_delimitada = img.crop(coordenadas[tipo])
  # convierte la imagen PIL a una arreglo numpy para poder usarlo con el modelo ocr
  delimitado = np.asarray(img_delimitada)
  result = modelo([delimitado])
  json = result.export()
  #display(img_delimitada)
  for pagina in json['pages']:
    for bloque in pagina['blocks']:
      for linea in bloque['lines']:
        linea=" ".join([w['value'] for w in linea['words']])
        arreglo.append(linea)
  if(tipo=="ticket"):
    for x in range(len(arreglo)): #['NO. DE TICKET FGR:', 'INC25368', 'NO. DE TICKET THEOS:', 'INC8942']
      if(re.search("FGR",arreglo[x])):
        v_fgr=arreglo[x+1]
      if(re.search("THEOS",arreglo[x])):
        v_theos=arreglo[x+1]
    if(v_fgr!="" and v_theos!=""):
      resultado=v_fgr+"-"+v_theos
      resultado=[resultado]
  if(tipo=="datos"):
    for x in range(len(arreglo)):
      if(re.search("USUARIO",arreglo[x])):
        if(re.search(":",arreglo[x])):
          nuevo=re.split(":",arreglo[x])
          v_nombre.append(nuevo[1])
      if(re.search("SEDE",arreglo[x])):
        v_cede.append(arreglo[x])
      if(re.search("SERIE",arreglo[x])):
        v_serie.append(arreglo[x+1])
      resultado=[v_nombre,v_cede,v_serie]
  return resultado

#Metodo para agregar nuevo renglon a la tabla
def nuevo_renglon(folio,usuario,ubicacion,serie): #def nuevo_renglon(folio,usuario,ubicacion,serie):
  renglon=tabla.add_row().cells
  renglon[0].text=folio
  renglon[1].text=usuario
  renglon[2].text=ubicacion
  renglon[3].text=serie
  #Formato del renglon
  for cell in renglon:
    parrafo=cell.paragraphs[0]
    for run in parrafo.runs:
      run.font.size=Pt(11)
      run.font.name="Calibri"
    parrafo.alignment=WD_ALIGN_PARAGRAPH.CENTER

def reconocimiento_resguardo(mes,resguardo):
  arreglo=[]
  #Lee imagen del acta

  coordenadas={'datos':[0,0,5100,3000],'ticket':[0,5800,5080,6250]}

  def reconocimiento(coordenada):
    doc = fitz.open(f"drive/MyDrive/PyFormatos/Archivos/{mes}2026/{resguardo}.pdf") # Lectura del PDF
    modelo=ocr_predictor(det_arch='db_resnet50',reco_arch='crnn_vgg16_bn',pretrained=True,assume_straight_pages=True)
    page = doc[0] # Seleccion de la hoja del archivo PDF
    zoom = 2  # Renderizar como imagen (controlas resolución aquí)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat,dpi=600)
    # Convertir a una imagen modelo RGB (3 Colores)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    # utiliza el metodo crop para delimitar la imagen (izquierda, arriba, derecha, abajo).

    img_delimitada = img.crop(coordenada)
    # convierte la imagen PIL a una arreglo numpy para poder usarlo con el modelo ocr
    delimitado = np.asarray(img_delimitada)
    result = modelo([delimitado])
    json = result.export()
    #display(img_delimitada)

    for pagina in json['pages']:
      for bloque in pagina['blocks']:
        for linea in bloque['lines']:
          linea=" ".join([w['value'] for w in linea['words']])
          arreglo.append(linea)
    return arreglo

  reconocimiento(coordenadas['datos'])
  datos=reconocimiento(coordenadas['ticket'])
  return datos

#Genera un nuevo resguardo -----------------------------------------------------
def resguardo(mes,archivo): #(mes,"resguardo-1") mes="junio"
  folio=[]
  id=[]
  serie=[]
  estado=[]
  usuario=[]
  #arr_datos=[]
  datos=reconocimiento_resguardo(mes,archivo) #(mes,"resguardo-1") mes="junio"
  global can_res
  for x in range(len(datos)):
    if(re.search("FOLIO",datos[x])): # ID, USUARIO, SERIE, ESTADO, TICKET, GUIA
      if(re.search(':',datos[x])):
        folio=re.split(':',datos[x])
        folio=folio[1]
        folio=re.sub(" ","",folio)
    if(re.search("ID:",datos[x])):
      id=re.split(":",datos[x])
      id=id[1]
      id=re.sub(" ","",id)
    if(re.search("SERIE",datos[x])):
      serie=datos[x+6]
    if(re.search("ESTADO:",datos[x])):
      estado=re.split(":",datos[x])
      estado=estado[1]
    if(re.search("INMUEBLE",datos[x])): #'INMUEBLE: SEDE TIJUANA'
      if(re.search(":",datos[x])):
        inmueble=re.split(":",datos[x])
        inmueble=inmueble[1]
    #'NOMBRE(S): DANIELA','APELLIDO PATERNO: ENRIQUEZ', 'APELLIDO MATERNO: CAMINO'
    if(re.search("NOMBRE",datos[x])):
      if(re.search(":",datos[x])):
        nombre=re.split(":",datos[x])
        usuario.append(nombre[1])
    if(re.search("PATERNO",datos[x])):
      if(re.search(":",datos[x])):
        apellido=re.split(":",datos[x])
        usuario.append(apellido[1])
    if(re.search("MATERNO",datos[x])):
      if(re.search(":",datos[x])):
        apellido=re.split(":",datos[x])
        usuario.append(apellido[1])
    if(re.search("TICKET",datos[x])):
      if(re.search(":",datos[x])):
        ticket=re.split(":",datos[x])
        ticket=ticket[1]
  if(len(usuario)==3):
    usuario="".join(usuario)

  datos_resguardo[can_res]={
      'FOLIO':folio,
      'ID':id,
      'SERIE':serie,
      'ESTADO':estado,
      'USUARIO':usuario,
      'TICKET': ticket
  }

  #Nuevo renglon docx
  nuevo_renglon(folio,usuario,inmueble,serie) #nuevo_renglon(folio,usuario,ubicacion,serie)

  #Generar excel de resguardo
  Excel_Resguardos(datos_resguardo[can_res])

  #Guardar cambios al docx
  documento.save(f"drive/MyDrive/PyFormatos/Nuevo/ACUSE BC,{mes}2026.docx")
  can_res+=1

  return datos_resguardo

