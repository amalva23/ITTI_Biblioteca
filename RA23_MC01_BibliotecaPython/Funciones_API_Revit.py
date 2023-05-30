# -*- coding: utf-8 -*-

###########################################################
# BIBLIOTECAS
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
import clr  # CommonLanguage Runtime
import sys

# Para trabajar con la RevitAPI
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

# Para trabajar con la RevitAPIUI
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter

# Para trabajar contra el documento y hacer transacciones
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Identificadores
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication  # interfaz
uidoc = uiapp.ActiveUIDocument  # Para que el usuario interactue con el doc.

# Otras bibliotecas
import System
from System.Collections.Generic import List  # Para generar iList
from System import DBNull
from System.IO import File
from System.Data import DataTable, DataColumn

# FUNCIONES
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Operaciones con listas
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

def a_lista(arg):
    """
    Uso: Evita errores con la entrada de inputs.
    Me convierte a lista si no lo era.
    """
    if hasattr(arg, '__iter__'):
        return arg
    else:
        return [arg]


def flatten(lista):
    """
    Uso: Aplanado de lista con multiples sub niveles.
    """
    salida = []
    for x in lista:
        if hasattr(x, "__iter__"):
            salida.extend(flatten(x))
        else:
            salida.append(x)
    return salida


def profundidad_lista(lista):
    """
    Uso: Pregunta el nivel de profundidad de una lista.
    Entrada: lista <list>: Lista.
    Salida: Numero.
    """
    funcion = lambda sublista: isinstance(sublista, list) and max(map(
        funcion, sublista)) + 1
    return funcion(lista)


def unique_items(lista):
    """
    Uso: Crea una lista con los elementos unicos, mantiene el orden.
    Entrada:
        lista <List>
    """
    claves = []
    for item in lista:
        if item not in claves:
            claves.append(item)
    return claves


def agrupar_por_clave(lista, indice=0):
    """
    Uso: Agrupa una lista por una clave en el indice especificado.
    Entrada:
        lista <List>
        indice <Integer>: indice de la clave en la lista.
    Salida: lista <List> agrupada por clave.
    """
    listaIndice = list(map(lambda x: x[indice], lista))
    values = unique_items(listaIndice)
    return [[y for y in lista if y[indice] == x] for x in values]


def dicc_agrupar_por_clave_indice(lista, indice=0):
    """
    Uso:
        Genera un diccionario. En la lista de claves hay multiples items
        repetidos, entonces los valores se agrupan.
    Entrada:
        claves <List[string]>
        valores <List>
    Salida:
        Diccionario con valores agrupados por clave <dict>
    """
    diccionario = dict()
    for a, b in zip(list(map(lambda x: x[indice], lista)), list(map(lambda x: x.pop(indice), lista))):
        if a in diccionario:
            diccionario[a].append(b)
        else:
            diccionario[a] = list()
            diccionario[a].append(b)
    return diccionario


def dicc_agrupar_por_clave(claves, valores):
    """
    Uso:
        Genera un diccionario. En la lista de claves hay multiples items
        repetidos, entonces los valores se agrupan.
    Entrada:
        claves <List[string]>
        valores <List>
    Salida:
        Diccionario con valores agrupados por clave <dict>
    """
    diccionario = dict()
    for a, b in zip(claves, valores):
        if a in diccionario:
            diccionario[a].append(b)
        else:
            diccionario[a] = list()
            diccionario[a].append(b)
    return diccionario


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Listados
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def listado_builtincategory():
    """
    Uso: Listado de las BuiltInCategory de Revit.
    """
    return System.Enum.GetValues(BuiltInCategory)


def listado_builtinparameter():
    """
    Uso: Listado de las BuiltInParameter de Revit.
    """
    return System.Enum.GetValues(BuiltInParameter)


def listado_categorias():
    """
    Uso: Genera un diccionario con todas las categorias de Revit.
    Salida: Diccionario(clave: nombre, valor: categoria)
    """
    diccionario = {}
    categorias = doc.Settings.Categories
    for cat in categorias:
        diccionario[cat.Name] = cat
    return diccionario


def listado_plantillas_vista(documento=doc):
    """
    Uso: Obtener el listado de todas las plantillas de vista.
    Entrada: Documento del que se obtiene el listado.
    Salida: Diccionario(clave: nombre plantilla, valor: id).
    """
    salida = dict()
    for vista in FilteredElementCollector(documento).OfClass(View):
        if vista.IsTemplate:
            salida[vista.Name] = vista.Id
    return salida


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Filtros
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def bic_por_nombrevisual(nombre):
    """
    Uso: Busca la BuiltInCategory que coincide con el nombre dado.
    Entrada: Nombre de la categoria a buscar <string>.
    Salida: BuiltInCategory <BuiltInCategory>.
    """
    values = System.Enum.GetValues(BuiltInCategory)

    def evitar_error(default, function, *args, **kwargs):
        """
        Uso: Sustituye un posible error devuelto por la función aplicada por None.
        """
        try:
            return function(*args, **kwargs)
        except:
            return default

    values_revit = map(lambda x: evitar_error(None, LabelUtils.GetLabelFor, x), values)
    dictionary = dict(zip(values_revit, values))
    return dictionary[nombre]


def bic_por_nombrebuilt(nombre):
    """
    Uso: Busca la BuiltInCategory que coincide con el nombre dado.
    Entrada: Nombre de la BuiltInCategory (a partir del punto) a buscar <string>.
    Salida: BuiltInCategory <BuiltInCategory>.
    """
    values = System.Enum.GetValues(BuiltInCategory)
    names = System.Enum.GetNames(BuiltInCategory)

    dictionary = dict(zip(names, values))
    return dictionary[nombre]


def bip_por_nombrevisual(nombre):
    """
    Uso: Busca el BuiltInParameter que coincide con el nombre dado.
    Entrada: Nombre del parametro a buscar <string>.
    Salida: BuiltInParameter <BuiltInParameter>.
    """
    values = System.Enum.GetValues(BuiltInParameter)

    def evitar_error(default, function, *args, **kwargs):
        """
        Uso: Sustituye un posible error devuelto por la función aplicada por None.
        """
        try:
            return function(*args, **kwargs)
        except:
            return default

    values_revit = map(lambda x: evitar_error(None, LabelUtils.GetLabelFor, x), values)
    dictionary = dict(zip(values_revit, values))
    return dictionary[nombre]


def bip_por_nombrebuilt(nombre):
    """
    Uso: Busca el BuiltInParameter que coincide con el nombre dado.
    Entrada: Nombre del BuiltInParameter (a partir del punto) a buscar <string>.
    Salida: BuiltInParameter <BuiltInParameter>.
    """
    values = System.Enum.GetValues(BuiltInParameter)
    names = System.Enum.GetNames(BuiltInParameter)

    dictionary = dict(zip(names, values))
    return dictionary[nombre]


def id_sharedparameter_nombre(nombre):
    """
    Uso: Obtiene un parametro de usuario por nombre.
    Entrada: nombre <str>: Nombre del parametro.
    Salida: Id del parametro <ElementId>.
    """
    par = None
    iterador = doc.ParameterBindings().ForwardIterator()
    while iterador.MoveNext():
        if iterador.Key.Name == nombre:
            par = iterador.Key.Id
            break
    return par


def filtro_parametro_pers(nombre, valor):
    """
    Uso: Crea un filtro para el FilteredElementCollector.
    Entrada:
        nombre <str>: Nombre del parametro
        valor <str>: Valor del parametro
    Salida: Filtro para aplicar en FilteredElementCollector con .WherePasses()
    """
    proveedor = None
    iterador = doc.ParameterBindings.ForwardIterator()
    while iterador.MoveNext():
        if iterador.Key.Name == nombre:
            proveedor = ParameterValueProvider(iterador.Key.Id)
            break
    evaluador = FilterStringEquals()
    regla = FilterStringRule(proveedor, evaluador, valor)
    return ElementParameterFilter(regla)


def filtro_parametro_bip(bip, valor):
    """
    Uso: Crea un filtro para el FilteredElementCollector.
    Entrada:
        nombre <str>: Nombre del parametro
        valor <str>: Valor del parametro
    Salida: Filtro para aplicar en FilteredElementCollector con .WherePasses()
    """
    proveedor = ParameterValueProvider(bip)
    evaluador = FilterStringEquals()
    regla = FilterStringRule(proveedor, evaluador, valor)
    return ElementParameterFilter(regla)


def filtrar_vista_por_nombre(nombre):
    """
    Uso:
    Entrada:
        nombre <str>: El nombre de la vista a filtrar
    Salida: La vista <View>
    """
    salida = dict()
    colector = FilteredElementCollector(doc).OfClass(View).ToElements()
    # Se limpian las plantillas y los nulos
    for vista in colector:
        if vista.IsTemplate is not True:
            salida[vista.Name] = vista
    # Seguro para evitar errores
    try:
        return salida[nombre]
    except:
        return 'Error: revisar nombre de entrada.'


def filtrar_nivel_por_nombre(nombre, doc = DocumentManager.Instance.CurrentDBDocument):
    """
    Uso:
    Entrada:
        nombre <str>: El nombre de la vista a filtrar
    Salida: La vista <View>
    """
    salida = dict()
    colector = FilteredElementCollector(doc).OfClass(Level).ToElements()
    # Se limpian las plantillas y los nulos
    for nivel in colector:
        salida[nivel.Name] = nivel
    # Seguro para evitar errores
    try:
        return salida[nombre]
    except:
        return 'Error: revisar nombre de entrada.'


def filtrar_vista_por_tipo(nombre):
    """
    Uso: Filtra las vistas por el tipo dado.
    Entrada: nombre <str>
    """
    names = System.Enum.GetNames(ViewType)
    values = System.Enum.GetValues(ViewType)

    tiposVistas = dict(zip(names, values))
    # Returns a BuiltInParameter given its name (key)
    if nombre in tiposVistas:
        vistas = FilteredElementCollector(doc).OfClass(View)
        salida = [v for v in vistas if v.ViewType == tiposVistas[nombre]]
    else:
        salida = "Revisar valor entrada"

    return salida


def filtrar_patron_por_nombre(nombre):
    """
    Uso: Selecciona el patron de relleno por nombre, si existe.
    Entrada:
        nombre <str>: El nombre del patron a seleccionar
    Salida: La vista <View>
    """
    salida = None
    colector = (
        FilteredElementCollector(doc).OfClass(FillPatternElement).
        ToElements())
    # Se revisa el nombre de los filtros
    for filtro in colector:
        if filtro.Name == nombre:
            salida = filtro
    return salida


def filtrar_esquema_por_nombre(nombre):
    """
    Uso: Selecciona el esquema de area por nombre, si existe.
    Entrada:
        nombre <str>: El nombre deL esquema a seleccionar
    Salida: La vista <View>
    """
    salida = None
    colector = (
        FilteredElementCollector(doc).OfClass(AreaScheme).
        ToElements())
    # Se revisa el nombre de los filtros
    for esquema in colector:
        if esquema.Name == nombre:
            salida = esquema
    return salida


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Utilidades
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def evitar_error(default, function, *args, **kwargs):
    """
    Uso:
        Atrapa errores convirtiendolos en Null. Para map() o list
        comprehension sin fallos.
    Entrada:
        default: Sustituto del error
        function: Funcion (sin parentesis)
        *args = Elemento sobre el que aplicar la funcion
    """
    try:
        return function(*args, **kwargs)
    except:
        return default


def parametro_valor(par):
    """
    Uso: Obtener el valor de cualquier parametro, sea cual sea su Storage.
    Entrada:
        par <Parameter>: Parametro a consultar su valor.
    Salida: El valor del parametro.
    """
    if par is not None:
        if par.StorageType == StorageType.String:
            valor = par.AsString()
        elif par.StorageType == StorageType.Double:
            valor = par.AsDouble()
        elif par.StorageType == StorageType.Integer:
            valor = par.AsInteger()
        elif par.StorageType == StorageType.ElementId:
            valor = doc.GetElement(par.AsElementId())
        else:
            valor = par.AsValueString()
    else:
        valor = None
    return valor


def valor_parametro_multiple(e, nombre):
    """
    Uso: Lee todos los parametros que coinciden con el nombre dado, y obtiene
    un valor que no sea nulo.
    Entrada: e <Element>
             nombre <string>
    """
    for par in e.GetParameters(nombre):
        if par.AsString() != "" and par.AsString() is not None:
            return par.AsString()


def valor_bip_reemplazavacios(e, bip):
    """
    Uso: Obtiene el valor del parametro, y lo reemplaza si sale null.
    Entrada: e <Element>
             bip <BuiltInParameter>
    """
    bip = e.get_Parameter(bip)
    if bip is None: return " "
    else: return bip.AsString()


def obtener_tipo(elemento):
    """
    Uso: Obtiene el tipo desde un elemento.
    """
    tipo = doc.GetElement(elemento.GetTypeId())
    return tipo


def color_consultar_rgb(color):
    """
    Uso: Consulta los valores rgb.
    Salida: Canales rgb en formato string.
    """
    if isinstance(color, Color):
        # Se revisa si el color es valido o no ha sido definido por el usuario
        if color.IsValid:
            salida = "(%s, %s, %s)" % (color.Red, color.Green, color.Blue)
        else:
            salida = -1
    else:
        salida = 'Se esperaba <Autodesk.Revit.DB.Color>'
    return salida


def natural_keys(texto):
    """
    Uso: Convierte un texto en letras y numeros para su orden natural.
    Entrada: texto <string>
    Salida: lista con caracteres del texto separados.
    """
    from re import split

    def atoi(texto1):
        """
        Uso: Convierte un texto numero si es numero.
        """
        return int(texto1) if texto1.isdigit() else texto1

    return [atoi(c) for c in split(r'(\d+)', texto)]


def elimina_tildes(texto):
    """
    Uso: Elimina las tildes de un texto.
    Entrada: texto <string>
    """
    import unicodedata
    s = ''.join((c for c in unicodedata.normalize('NFD', texto)
                 if unicodedata.category(c) != 'Mn'))
    return s


def sensible_title_caps(texto, no_caps_list=('y', 'o', 'e', 'de')):
    """
    Uso: Pone en mayusculas la primera letra de cada palabra. Si no reconoce algún signo,
    añadirlo en la funcion split() separado por '|'.
    """
    words = []
    import re
    for word in re.split(r'(\W|_)', texto):
        if word not in no_caps_list:
            word = word.capitalize()
        words.append(word)
    return "".join(words)


class CreateFailureAdvancedHandler(IFailuresPreprocessor):
    """
    Uso: (API) An interface that may be used to perform a preprocessing
        step to either filter out anticipated transaction failures or to mark certain failures as non-continuable.
        An instance of this interface can be set in the failure handling options of transaction object.
    """
    def __init__(self):
        self.ErrorMessage = []

    def PreprocessFailures(self, failuresAccessor):
        # Inside event handler, get all warnings
        failList = failuresAccessor.GetFailureMessages()

        if failList.Count == 0:
            return FailureProcessingResult.Continue
        else:
            for failure in failList:
                iD = failure.GetFailureDefinitionId()
                failureSeverity = failure.GetSeverity()
                if failureSeverity == FailureSeverity.Warning:
                    try:
                        errormsg = failure.GetDescriptionText()
                    except:
                        errormsg = "Error desconocido"
                        # Si el aviso es de instancias identicas en el mismo lugar, rollback
                    if BuiltInFailures.OverlapFailures.DuplicateInstances == iD:
                        return FailureProcessingResult.ProceedWithRollBack
                        # Si el aviso es de valor de marca duplicado, ignora
                    elif BuiltInFailures.GeneralFailures.DuplicateValue == iD:
                        errormsg = "Elementos tienen valores duplicados"
                        failuresAccessor.DeleteWarning(failure)

                    self.ErrorMessage.append(errormsg)

                    return FailureProcessingResult.Continue


def get_selection():
    ids = uidoc.Selection.GetElementIds()
    elems = []
    for id in ids:
        try:
            elems.append(doc.GetElement(id))
        except:
            elems.append('Fallo al leer la seleccion.')
    return elems


def select_linked_elements():
    refs = uidoc.Selection.PickObjects(Selection.ObjectType.LinkedElement, "pick an element in the linked model")
    refs = list(refs)
    elems = []
    for ref in refs:
        linkedDoc = doc.GetElement(ref.ElementId).GetLinkDocument()
        try:
            resultado = linkedDoc.GetElement(ref.LinkedElementId)
        except:
            resultado = 'Fallo al leer la seleccion.'
        elems.append(resultado)
    return elems


def rango(inicio, fin, paso):
    """
    Uso: Crea una lista de numeros con el rango especificado.
    Entrada:
        inicio: Numero incial
        fin = Numero final
        paso = espacio entre numeros
    Salida: lista de numeros
    """
    salida = []
    while inicio <= fin:
        salida.append(inicio)
        inicio += paso
    return salida


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Utilidades VISTAS
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def vista_aplicar_configuracion_deotra(vista, configuracion):
    """
    Uso: Aplica la configuracion de la vista configuracion.
    Entrada:
        vista <View>
        configuracion <View>
    """
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        vista.ApplyViewTemplateParameters(configuracion)
        TransactionManager.Instance.TransactionTaskDone()
        return 'Completado'
    except:
        return 'Fallo: revisar tipos de vista.'


def vista_definir_vista_activa(vista):
    """
    Uso: Forzar que una vista en concreto sea la vista activa.
    Entrada: vista <View>
    """
    TransactionManager.Instance.ForceCloseTransaction()
    try:
        uidoc.RequestViewChange(vista)
        return 'Completado'
    except:
        return 'Fallo'


def id_tipo_familia_vista():
    """
    Uso: Diccionario de tipos de vista. Devuelve el id del tipo.
         Ejemplo: id_tipo_familia_vista()['Tipo de vista'] = idVista
    """
    colector = (FilteredElementCollector(doc).OfClass(ViewFamilyType).
                ToElements())
    ids = [tipo.Id for tipo in colector]
    nombres = [str(tipo.ViewFamily) for tipo in colector]
    return dict(zip(nombres, ids))


def limpieza_vistas_sin_uso(inicio, prefijo='', opcion_tablas=True):
    """
    Uso: Se eliminan las vistas sin uso, teniendo en cuenta anfitrio de vistas.
    dependientes, de llamada...
    Entrada: inicio <boolean>
    Salida: Resultado
    """
    if inicio:
        # Colectar los planos
        planos = FilteredElementCollector(doc).OfClass(ViewSheet)
        # Se consultan las vistas dentro de los planos
        idsvistasPlanos = [p.GetAllPlacedViews() for p in planos]
        # Aplano
        idsUsados = [v for lista in idsvistasPlanos for v in lista]

        # Reviso la dependencia
        idsDepen = [doc.GetElement(v).GetPrimaryViewId() for v in idsUsados]
        for idNum in idsDepen:
            if str(idNum) != '-1' and idNum not in idsUsados:
                idsUsados.append(idNum)

        # Reviso callouts
        idsCalloutanfi = [doc.GetElement(idNum).get_Parameter(BuiltInParameter.SECTION_PARENT_VIEW_NAME) for idNum in
                          idsUsados]
        for parametro in idsCalloutanfi:
            if parametro is not None:
                idNum = parametro.AsElementId()
                if idNum != ElementId(-1) and idNum not in idsUsados:
                    idsUsados.append(parametro.AsElementId())

        # Se incorporan los planos
        for plano in planos:
            idsUsados.append(plano.Id)

        # Se trabaja con las tablas
        if opcion_tablas:
            # Se conservan las tablas
            tablas = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
            if tablas:
                for t in tablas:
                    idsUsados.append(t.Id)
        else:
            # Se quiere eliminar todas las tablas sin uso
            tablas = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
            if tablas:
                for t in tablas:
                    idsUsados.append(t.ScheduleId)

                    # Colectar todas las vistas
        vistas = FilteredElementCollector(doc).OfClass(View).ToElements()
        # Descartar plantillas de vistas
        vistasIds = []
        for v in vistas:
            if v.IsTemplate is False:
                if v.Name.startswith(prefijo) is False \
                        and v.ViewType != ViewType.SystemBrowser and v.ViewType != ViewType.ProjectBrowser:
                    vistasIds.append(v.Id)

        # Resto la lista que quiero mantener a la de vistas a eliminar
        lista1 = set(idsUsados)
        lista2 = set(vistasIds)
        idsNoUsados = lista2.difference(lista1)

        contador = 0
        if idsNoUsados:
            TransactionManager.Instance.EnsureInTransaction(doc)
            for idNum in idsNoUsados:
                try:
                    doc.Delete(idNum)
                    contador += 1
                except Exception:
                    raise

            TransactionManager.Instance.TransactionTaskDone()

        if contador == 0 and len(idsNoUsados) != 0:
            salida = 'Existen {} vistas sin usar y no se ha eliminado ninguna.'.format(len(idsNoUsados))
        elif contador > 0:
            salida = 'Se han eliminado {} vistas de {} vistas posibles.'.format(contador, len(idsNoUsados))
        else:
            salida = 'No hay vistas sin uso'

    else:
        salida = 'Aviso: Para iniciar la ejecución necesita un True.'

    return salida


def limpieza_plantillas_vista(inicio):
    """
    Uso: Se eliminan las plantillas sin uso.
    """
    if inicio:
        plantillas, plantillasUso = set(), set()
        for vista in FilteredElementCollector(doc).OfClass(View).ToElements():
            if vista.IsTemplate:
                plantillas.add(vista.Id)
            else:
                if vista.ViewTemplateId != ElementId(-1):
                    plantillasUso.add(vista.ViewTemplateId)

        # Hacemos la diferencia de conjuntos
        sinUso = plantillas.difference(plantillasUso)
        limpieza = List[ElementId](sinUso)
        # Revisamos el contenido de sinUso
        if limpieza:
            try:
                TransactionManager.Instance.EnsureInTransaction(doc)
                doc.Delete(limpieza)
                TransactionManager.Instance.TransactionTaskDone()
                salida = 'Completado'
            except:
                salida = 'Fallo: En la eliminación.'
        else:
            salida = 'Completado.\nNo existen plantillas sin uso.'
    else:
        salida = 'Aviso: Para iniciar la ejecución necesita un True.'

    return salida


def vista_crear_plantilla(lista):
    """
    Uso:
        Crear, una o varias, plantilla/s de vista partiendo de una
        o multiples vistas.
    Entrada:
        lista <>: Desconocemos si nos dan una vista o una lista de vistas.
    Salida: Plantilla/s de vista generada/s.
    Cuidado: Dependencia de la funcion aLista().
    """
    vistas = a_lista(lista)
    TransactionManager.Instance.EnsureInTransaction(doc)
    salida = [vista.CreateViewTemplate() for vista in vistas]
    TransactionManager.Instance.TransactionTaskDone()
    return salida


def graficos_aislar_elementos_temporal(lista, vista=doc.ActiveView):
    """
    Uso: 
    Entrada:
        lista <>: Desconocemos si nos daran una vista o una lista de vistas
        vista <View>: valor por defecto la vista activa
    Salida: Mensaje exito/fallo.
    Cuidado: Dependencia de la funcion aLista().
    """
    elementos = a_lista(lista)
    idsLista = List[ElementId]()
    for e in elementos:
        idsLista.Add(e.Id)
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        vista.IsolateElementsTemporary(idsLista)
        salida = 'Completado'
        TransactionManager.Instance.TransactionTaskDone()
    except:
        salida = None
    return salida


def graficos_aislar_elementos(lista, vista=doc.ActiveView):
    """
    Uso:
    Entrada:
        lista <>: Desconocemos si nos daran una vista o una lista de vistas
        vista <View>: defecto vista activa
    Salida: Mensaje exito/fallo.
    Cuidado: Dependencia de la funcion aLista().
    """
    elementos = a_lista(lista)
    idsLista = List[ElementId]()
    for e in elementos:
        idsLista.Add(e.Id)
    try:
        TransactionManager.Instance.EnsureInTransaction(doc)
        vista.IsolateElementsTemporary(idsLista)
        vista.ConvertTemporaryHideIsolateToPermanent()
        salida = 'Completado'
        TransactionManager.Instance.TransactionTaskDone()
    except:
        salida = None
    return salida


def vista_opciones_duplicado(integer):
    """
    Uso: Selecciona un ViewDuplicateOption.
         1 = Duplicate   2 = AsDependent   3 = WithDetailing
    """
    opciones = System.Enum.GetValues(ViewDuplicateOption)

    if integer <= 2:
        return opciones[integer]
    else:
        return 'Fallo: Introducir un valor entre 0 y 2.'


def vista_duplicar(lista, opciones=1, prefijo='', sufijo=''):
    """
    Uso:
    Entrada:
        lista <>: Desconocemos si nos daran una vista o una lista de vistas
        opciones <int>: Opciones de duplicado, defecto 1
        prefijo <str>: Prefijo para el nombre de la vista, defecto ''
        sufijo <str>: Prefijo para el nombre de la vista, defecto ''
    Salida: Mensaje exito/fallo.
    Cuidado: Dependencia de la funcion aLista() y vista_opciones_duplicado().    
    """
    vistas = a_lista(lista)

    salida = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    for vista in vistas:
        if not vista.IsTemplate:
            idNum = vista.Duplicate(vista_opciones_duplicado(opciones))
            nueva = doc.GetElement(idNum)
            try:
                nueva.Name = str(prefijo) + vista.Name + str(sufijo)
            except:
                pass
            salida.append(nueva)
    TransactionManager.Instance.TransactionTaskDone()

    return salida


def vista_consultar_modificaciones_filtros(vista):
    """
    Uso:
        Se revisa la vista y se extrae la informacion sobre los filtros y que
        modificaciones hay en la visibilidad grafica.
    Entradas:
        vista <View>: Vista a revisar
    Salida:
        Se genera un diccionario con diccionarios anidados, para acceder por
        clave a cualquier configuracion dentro de cualquier filtro.
    """
    salida = dict()
    if isinstance(vista, View) and vista.IsTemplate is False:
        try:
            # Se buscan los filtros
            filtros = vista.GetFilters()
            # Se condiciona el siguiente paso: debe haber filtros
            if filtros:
                for idNum in filtros:
                    datos = dict()
                    # Se consulta el nombre
                    # Sera la clave para acceder al filtro
                    filtro = doc.GetElement(idNum)
                    nombreFiltro = filtro.Name
                    # Se consulta si esta visible y activo el filtro
                    datos['Visible'] = vista.GetFilterVisibility(idNum)
                    datos['Activo'] = vista.GetIsFilterEnabled(idNum)

                    # Se accede a las modificaciones
                    configuracion = vista.GetFilterOverrides(idNum)
                    # Se va a generar un sistema de diccionarios anidados

                    # Informacion de proyeccion
                    proyeccion = dict()
                    # Informacion lineas
                    lineasProyeccion = dict()
                    lineasProyeccion['Patrón'] = (
                        configuracion.ProjectionLinePatternId)
                    lineasProyeccion['Color'] = (
                        color_consultar_rgb(
                            configuracion.ProjectionLineColor))
                    lineasProyeccion['Grosor'] = (
                        configuracion.ProjectionLineWeight)
                    # Se almacena la informacion lineas de proyeccion
                    proyeccion['Líneas'] = lineasProyeccion
                    # Informacion de los patrones de superficie
                    patronesProyeccion = dict()
                    # Primero lo relativo al patron primer plano
                    patronesProyeccion['Primer plano - Visibilidad'] = (
                        configuracion.IsSurfaceForegroundPatternVisible)
                    patronesProyeccion['Primer plano - Patrón'] = (
                        configuracion.SurfaceForegroundPatternId)
                    patronesProyeccion['Primer plano - Color'] = (
                        color_consultar_rgb(
                            configuracion.SurfaceForegroundPatternColor))
                    # Lo relativo al patron de fondo
                    patronesProyeccion['Fondo - Visibilidad'] = (
                        configuracion.IsCutForegroundPatternVisible)
                    patronesProyeccion['Fondo - Patrón'] = (
                        configuracion.CutForegroundPatternId)
                    patronesProyeccion['Fondo - Color'] = (
                        color_consultar_rgb(
                            configuracion.CutForegroundPatternColor))
                    # Se almacena la informacion lineas de proyeccion
                    proyeccion['Patrones'] = patronesProyeccion
                    # Tercero: transparencia
                    proyeccion['Transparencia'] = configuracion.Transparency
                    # Se almacena toda la informacion relativa a proyeccion
                    datos['Proyeccion'] = proyeccion

                    # Informacion de corte
                    corte = dict()
                    # Informacion lineas
                    lineasCorte = dict()
                    lineasCorte['Patrón'] = (
                        configuracion.CutLinePatternId)
                    lineasCorte['Color'] = (
                        color_consultar_rgb(configuracion.CutLineColor))
                    lineasCorte['Grosor'] = (
                        configuracion.CutLineWeight)
                    # Se almacena la informacion lineas de proyeccion
                    corte['Líneas'] = lineasCorte
                    # Informacion de los patrones de superficie
                    patronesCorte = dict()
                    # Primero lo relativo al patron primer plano
                    patronesCorte['Primer plano - Visibilidad'] = (
                        configuracion.IsSurfaceBackgroundPatternVisible)
                    patronesCorte['Primer plano - Patrón'] = (
                        configuracion.SurfaceBackgroundPatternId)
                    patronesCorte['Primer plano - Color'] = (
                        color_consultar_rgb(
                            configuracion.SurfaceBackgroundPatternColor))
                    # Lo relativo al patron de fondo
                    patronesCorte['Fondo - Visibilidad'] = (
                        configuracion.IsCutBackgroundPatternVisible)
                    patronesCorte['Fondo - Patrón'] = (
                        configuracion.CutBackgroundPatternId)
                    patronesCorte['Fondo - Color'] = (
                        color_consultar_rgb(
                            configuracion.CutBackgroundPatternColor))
                    # Se almacena la informacion lineas de proyeccion
                    corte['Patrones'] = patronesCorte
                    # Se almacena toda la informacion relativa a corte
                    datos['Corte'] = corte

                    # Se consulta si esta a medio tono
                    datos['Tramado'] = configuracion.Halftone
                    salida[nombreFiltro] = datos
            else:
                salida = 'No hay filtros aplicados en la vista.'
        except:
            salida = 'No se puede aplicar \nfiltros a esta vista.'
    else:
        salida = ('Se esperaba una vista.\n'
                  'Revisar si es una plantilla de vista.')
    return salida


def seccion_paralela_por_curva(curva, desfase, altura):
    """
    Uso:
    Entrada:
        curva <curve>: Revit curve.
        desfase <float>: Separacion del muro. Siempre en unidades metricas.
        altura <float>: Altura del muro.
    Salida:
    """
    # Convierto unidades a internas
    o = unidades_modelo_a_internas_longitud(float(desfase))
    h = unidades_modelo_a_internas_longitud(float(altura))
    # Revisar el tipo de curva
    # Si es una curva de Dynamo, la pasamos a Revit
    if isinstance(curva, Autodesk.DesignScript.Geometry.Line):
        c = curva.ToRevitType()
    else:
        c = curva

    # Obtener los puntos
    i = c.GetEndPoint(0)  # XYZ
    f = c.GetEndPoint(1)  # XYZ
    d = f - i  # XYZ
    long = d.GetLength()

    # Definir los ejes
    x = d.Normalize()
    y = XYZ.BasisZ
    z = x.CrossProduct(y)

    # Crear una instancia de la clase Transform
    t = Transform.Identity
    t.Origin = i + 0.5 * d
    t.BasisX = x
    t.BasisY = y
    t.BasisZ = z

    # Creamos BoundingBox
    caja = BoundingBoxXYZ()
    caja.Transform = t
    caja.Min = XYZ(- (0.5 * long + o), i.Z - o, - o)
    caja.Max = XYZ(0.5 * long + o, h + o, o)

    idTipoFamilia = id_tipo_familia_vista()['Section']

    TransactionManager.Instance.EnsureInTransaction(doc)
    seccion = ViewSection.CreateSection(doc, idTipoFamilia, caja)
    TransactionManager.Instance.TransactionTaskDone()

    return seccion


def seccion_perpendicular_por_curva(curva, desfase, altura):
    """
    Uso:
    Entrada:
        curva <curve>: Revit curve.
        desfase <float>: Separacion del muro. Siempre en unidades metricas.
        altura <float>: Altura del muro.
    Salida:
    """
    # Convierto unidades a internas
    o = unidades_modelo_a_internas_longitud(float(desfase))
    h = unidades_modelo_a_internas_longitud(float(altura))

    # Revisar el tipo de curva
    # Si es una curva de Dynamo, la pasamos a Revit
    if isinstance(curva, Autodesk.DesignScript.Geometry.Line):
        c = curva.ToRevitType()
    else:
        c = curva

    i = c.GetEndPoint(0)  # XYZ
    f = c.GetEndPoint(1)  # XYZ
    d = f - i  # XYZ
    # l = d.GetLength()  # Obtener la longiud

    # Definir los ejes
    z = d.Normalize()
    y = XYZ.BasisZ
    x = y.CrossProduct(z)

    # Crear una instancia de la clase Transform
    t = Transform.Identity
    t.Origin = i + 0.5 * d
    t.BasisX = x
    t.BasisY = y
    t.BasisZ = z

    # Creamos BoundingBox
    caja = BoundingBoxXYZ()
    caja.Transform = t
    caja.Min = XYZ(- o, i.Z - o, - o)
    caja.Max = XYZ(o, h + o, o)

    idTipoFamilia = id_tipo_familia_vista()['Section']

    TransactionManager.Instance.EnsureInTransaction(doc)
    seccion = ViewSection.CreateSection(doc, idTipoFamilia, caja)
    TransactionManager.Instance.TransactionTaskDone()

    return seccion


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Conversion de unidades
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def unidades_internas_a_modelo_longitud(valor):
    """
    Uso: Convierte unidades internas de longitud a unidades de modelo
    """
    if int(doc.Application.VersionNumber) >= 2022:
        unidadModelo = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
        return UnitUtils.ConvertFromInternalUnits(valor, unidadModelo)
    else:
        unidadModelo = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
        return UnitUtils.ConvertFromInternalUnits(valor, unidadModelo)


def unidades_internas_a_modelo_area(valor):
    """
    Uso: Convierte unidades internas de area a unidades de modelo
    """
    if int(doc.Application.VersionNumber) >= 2022:
        unidadModelo = doc.GetUnits().GetFormatOptions(SpecTypeId.Area).GetUnitTypeId()
        return UnitUtils.ConvertFromInternalUnits(valor, unidadModelo)
    else:
        unidadModelo = doc.GetUnits().GetFormatOptions(UnitType.UT_Area).DisplayUnits
        return UnitUtils.ConvertFromInternalUnits(valor, unidadModelo)


def unidades_modelo_a_internas_longitud(valor):
    """
    Uso: Convierte unidades de modelo de longitud a unidades internas
    """
    if int(doc.Application.VersionNumber) >= 2022:
        unidadModelo = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
        return UnitUtils.ConvertToInternalUnits(valor, unidadModelo)
    else:
        unidadModelo = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
        return UnitUtils.ConvertToInternalUnits(valor, unidadModelo)


def cm_a_internas_longitud(valor):
    """
    Uso: Convierte cm a unidades internas
    """
    if int(doc.Application.VersionNumber) >= 2022:
        return UnitUtils.ConvertToInternalUnits(valor, UnitType.Centimeters)
    else:
        return UnitUtils.ConvertToInternalUnits(valor, DisplayUnitType.DUT_CENTIMETERS)


def convertir_xyz_unidad_interna(punto):
    """
    Uso: Crea un XYZ() en metros, y lo convierte a unidades internas.
    Cuidado: Dependencia de la funcion unidades_modelo_a_internas_longitud().
    """
    factor = unidades_modelo_a_internas_longitud(1)
    return punto.Multiply(factor)


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Geometria
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
def group_curves(curves):
    """
    Uso: Agrupa las curvas continuas
    Entrada:
        curvas <list>
    """
    ignore_distance = 0.02  # Assume points this close or closer to each other are touching
    Grouped_Lines = []
    Queue = set()
    while curves:
        Shape = []
        Queue.add(curves.pop())  # Move a line from the curves to our queue
        while Queue:
            Current_Line = Queue.pop()
            Shape.append(Current_Line)
            for Potential_Match in curves:
                Points = (Potential_Match.GetEndPoint(0), Potential_Match.GetEndPoint(1))
                for P1 in Points:
                    for P2 in (Current_Line.GetEndPoint(0), Current_Line.GetEndPoint(1)):
                        distance = P1.DistanceTo(P2)
                        if distance <= ignore_distance:
                            Queue.add(Potential_Match)
            curves = [item for item in curves if item not in Queue]
        Grouped_Lines.append(Shape)
    return Grouped_Lines


def sort_curves_contiguous(curves):
    """
    Uso: Ordena un grupo de curvas, (el final de la 1º es igual al inicio de la 2º).
    Salida: Aplica la funcion sobre la lista original, no devuelve nada.
    """
    n = curves.Count

    # Walk through each curve (after the first)
    # to match up the curves in order
    found = 0
    for i in range(0, n, 1):
        curve = curves[i]
        endPoint = curve.GetEndPoint(1)

        # Find curve with start point = end point
        found = (i + 1 >= n)

        for j in range(i + 1, n, 1):
            k = j
            p = curves[k].GetEndPoint(0)

            # If there is a match end->start,
            # this is the next curve
            if 0.1 > p.DistanceTo(endPoint):
                if i + 1 != k:
                    tmp = curves[i + 1]
                    curves[i + 1] = curves[k]
                    curves[k] = tmp
                found = True
                break

            p = curves[k].GetEndPoint(1)

            # If there is a match end->end,
            # reverse the next curve
            if 0.1 > p.DistanceTo(endPoint):
                if i + 1 == k:
                    curves[i + 1] = curves[k].CreateReversed()
                else:
                    tmp = curves[i + 1]
                    curves[i + 1] = curves[k].CreateReversed()
                    curves[k] = tmp
                found = True
                break
        if not found == 0:
            Exception("SortCurvesContiguous:" + " non-contiguous input curves")


def point_in_polygon(poligono, pt):
    """
    Uso: testea si el punto 2D esta dentro del poligono 2D (no tiene en cuenta las Z)
    Entrada:
        polygon <List>: lista de curvas <Curve> ORDEN CLOCKWISE
        point <XYZ>: punto a testear.
    Salida: boolean
    """
    pt = [pt.X, pt.Y]
    polygon = [[x.GetEndPoint(0).X, x.GetEndPoint(0).Y] for x in poligono]
    odd = False
    # For each edge (In this case for each point of the polygon and the previous one)
    i = 0
    j = len(polygon) - 1
    while i < len(polygon):
        # If a line from the point into infinity crosses this edge
        # One point needs to be above, one below our y coordinate

        if (((polygon[i][1] > pt[1]) != (polygon[j][1] > pt[1])) and (pt[0] < (
                (polygon[j][0] - polygon[i][0]) * (pt[1] - polygon[i][1]) / (polygon[j][1] - polygon[i][1])) +
                                                                      polygon[i][0])):
            # Invert odd
            odd = not odd
        j = i
        i += 1
    # If the number of crossings was odd, the point is in the polygon
    return odd


def punto_xyz(x=0, y=0, z=0):
    """
    Uso: Crea un XYZ() en metros, y lo convierte a unidades internas.
    Cuidado: Dependencia de la funcion unidades_modelo_a_internas_longitud().
    """
    factor = unidades_modelo_a_internas_longitud(1)
    return XYZ(x, y, z).Multiply(factor)


def windows_search_files(ext, ruta, sub=True, bkup=True):
    """
    Uso: Busca archivos con la extension elegida en una ruta.
    Entrada: bkup<bool>: Incluye backups.
             sub<bool>: Incluye subcarpetas.
    Salida: lista de archivos<list>
    """
    import os
    salida = []
    if sub:
        for ruta, carpeta, archivos in os.walk(ruta):
            for archivo in archivos:
                if bkup is True and archivo.endswith(ext):
                    rutaArchivo = os.path.join(ruta, archivo)
                    salida.append(rutaArchivo)
                elif (bkup is False and archivo.endswith(
                        ext) and '.00' not in archivo):
                    rutaArchivo = os.path.join(ruta, archivo)
                    salida.append(rutaArchivo)
    else:
        for item in os.listdir(ruta):
            if bkup is True and item.endswith(ext):
                rutaArchivo = os.path.join(ruta, item)
                salida.append(rutaArchivo)
            elif bkup is False and item.endswith(
                    ext) and '.00' in item:
                rutaArchivo = os.path.join(ruta, item)
                salida.append(rutaArchivo)

                # for archivo in os.listdir(ruta):
    return salida


def round_nearest(x, a):
    """
    Uso: Redondea el numero a los decimales que quieras.
    Entrada: x <float>: numero inicial
             a <float>: decimal
    """
    return round(round(x / a) * a, 2)


def round_up(x, a):
    """
    Uso: Redondea el numero por arriba a los decimales que quieras.
    Entrada: x <float>: numero inicial
             a <float>: decimal
    """
    return round(math.ceil(x / a) * a, 2)


# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Clases
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
class MySelectionFilter(ISelectionFilter):
    def __init__(self):
        pass
    def AllowElement(self, element):
        bic = IN[0][0][0]
        if element.Category.Id == Category.GetCategory(doc,bic).Id:
            return True
        else:
            return False
    def AllowReference(self, element):
        return False


# Operaciones con archivos externos
def fill_list_value(list, lenght, value):
    return list + [value] * (lenght - len(list))


def get_data_txt(ruta):
    f = File.ReadAllLines(ruta)
    limpieza = [x for x in f if not x.startswith("#") and not x.startswith("*") and x]
    datos = list(map(lambda x: fill_list_value(x.split('	'), 10, ""), limpieza))
    return datos


def array_to_datatable(datos):
    tabla = DataTable()
    encabezado = datos.pop(0)
    for i in range(len(encabezado)):
        column = DataColumn(encabezado[i], System.Type.GetType("System.String"))
        column.DefaultValue = ""
        tabla.Columns.Add(column)

    for i in range(len(datos)):
        if sys.version_info[0] >= 3:
            fila = datos[i]
        else:
            fila = Array[str](datos[i])
        tabla.Rows.Add(fila)
    return tabla


def datatable_to_string(datatable, textoinicial):
    texto = ["	".join(list(map(lambda y: "" if type(y) == DBNull else y, x.ItemArray))) for x in datatable.Rows]
    texto.insert(0, "	".join([x.ColumnName for x in datatable.Columns]))
    texto.insert(0, textoinicial)
    return texto