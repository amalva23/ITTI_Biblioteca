# -*- coding: utf-8 -*-

import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import os.path
from System import Single, Int32, String

import System.Windows.Forms as wforms

from System.Windows.Forms import Form, \
    ListView, View, HorizontalAlignment, Button, MessageBox, MessageBoxButtons, DialogResult, \
    MessageBoxIcon, DataGridView

from System.Drawing import Size, Color, Icon, Point, GraphicsUnit, Font, FontFamily, \
    FontStyle, SizeF

from System.IO import File


from RA23_MC01_BibliotecaPython.Funciones_API_Revit import get_data_txt, array_to_datatable, \
    datatable_to_string

CURR_DIR = os.path.dirname(__file__)


def myFont(name, size, fontstyle=FontStyle.Regular):
    """'Crear una fuente"""
    font = Font.Overloads[FontFamily, Single, FontStyle, GraphicsUnit](FontFamily(name), size,
                                                                       fontstyle,
                                                                       GraphicsUnit.Pixel)
    return font


class MyWindow(Form):
    """Clase para crear interfaces graficas"""

    def __init__(self, title, dMaxX, dMaxY, dMinX, dMinY, opacity, transparency=0.9,
                 color=Color.WhiteSmoke,
                 icon=os.path.join(CURR_DIR, r"Logo\wall.ico"),
                 image=os.path.join(CURR_DIR, r"Logo\ITTI_Logo.png"),
                 pIX=None, pIY=None, dIX=None, dIY=None):
        """Constructor de la clase MyWindow"""

        self.Text = title

        if dMinX == dMaxX and dMinY == dMaxY:
            # En visual studio, para ver a tamano real de ejecucion en dynamo
            # self.AutoScaleDimensions = SizeF(96F, 96F)
            # self.AutoScaleMode = wforms.AutoScaleMode.Dpi
            self.ClientSize = Size(dMaxX, dMaxY)
            self.FormBorderStyle = wforms.FormBorderStyle.FixedDialog
        else:
            self.FormBorderStyle = wforms.FormBorderStyle.Sizable
            self.MinimumSize = Size(dMinX, dMinY)
            self.MaximumSize = Size(dMaxX, dMaxY)

        if opacity:
            self.AllowTransparency = opacity
            self.Opacity = transparency

        if icon:
            self.Icon = Icon(icon)

        if image:
            self.imagen = wforms.PictureBox()
            self.imagen.Load(image)
            self.imagen.SizeMode = wforms.PictureBoxSizeMode.StretchImage
            self.imagen.ClientSize = Size(dIX, dIY)
            self.imagen.Location = Point(pIX, pIY)
            self.Controls.Add(self.imagen)

        self.CenterToScreen()
        self.BringToFront()

        self.controlInfo = {
            'TextBox': {},
            'ComboBox': {},
            'CheckBox': {},
            'RadioButton': {},
            'ListView': {},
            'DataGridView': {},
        }

        self.dataGridView1 = DataGridView()
        self.checkboxCount = None
        self.checkboxName = None

    def run(self):
        """Muestra la interfaz en la ventana"""
        wforms.Application.Run(self)

    def myCheckBox(self, name, pX, pY, font, text):
        """Construir un checkBox"""

        self.checkBox = wforms.CheckBox()
        self.checkBox.Name = name
        self.checkBox.Location = Point(pX, pY)
        self.checkBox.Font = font
        self.checkBox.Text = text
        self.checkBox.AutoSize = True
        self.Controls.Add(self.checkBox)
        self.checkBox.Checked = True
        return self.checkBox

    def myCheckBoxGroup(self, pX, pY, dX, dY, cbpX, cbpY, gFont, cFont, text, elements, max=2):
        """Construir un grupo de check boxes"""

        checkBoxes = []
        Iactual = cbpY
        for element in elements:
            cb = self.myCheckBox('-'.join([text, element]), cbpX, Iactual, cFont, element)
            Iactual += cbpY
            checkBoxes.append(cb)

        groupBox = MyGroupBox(self, pX, pY, dX, dY, gFont, text, checkBoxes)
        self.checkboxCount = max
        self.checkboxName = text
        return groupBox

    def myTextBox(self, name, pX, pY, dX, dY, font, text):
        """Construir un textBox para grupos"""

        self.textBox = wforms.TextBox()
        self.textBox.Name = name
        self.textBox.Location = Point(pX, pY)
        self.textBox.ClientSize = Size(dX, dY)
        self.textBox.Font = font
        self.textBox.Text = text
        self.Controls.Add(self.textBox)

        return self.textBox

    def myTextBoxGroup(self, pX, pY, dX, dY, cbpX, cbpY, gFont, cFont, text, elements, tbX=200):
        """Construir un grupo de text boxes"""

        textBoxes = []
        Iactual = cbpY

        cbtX = 9*len(max(elements.keys(), key=lambda x: len(x)))
        for ctrlName in elements.keys():
            MyLabel(self, pX + cbpX, pY + Iactual, cFont, ctrlName + ":")
            tb = self.myTextBox('-'.join([text, ctrlName]), cbpX + cbtX, Iactual, tbX, 50, cFont,
                                elements[ctrlName])
            Iactual += cbpY
            textBoxes.append(tb)

        groupBox = MyGroupBox(self, pX, pY, dX, dY, gFont, text, textBoxes)
        return groupBox

    def myRadioButton(self, name, pX, pY, font, text):
        """Construir un checkBox"""

        self.radioButton = wforms.RadioButton()
        self.radioButton.Name = name
        self.radioButton.Location = Point(pX, pY)
        self.radioButton.Font = font
        self.radioButton.Text = text
        self.radioButton.AutoSize = True
        self.Controls.Add(self.radioButton)

        return self.radioButton

    def myRadioButtonGroup(self, pX, pY, dX, dY, rbpX, rbpY, gFont, rFont, text, elements):
        """Construir un grupo de check boxes"""

        radioButton = []
        Iactual = 20
        for element in elements:
            rb = self.myRadioButton('-'.join([text, element]), rbpX, Iactual, rFont, element)
            Iactual += rbpY
            radioButton.append(rb)
        radioButton[0].Checked = True
        groupBox = MyGroupBox(self, pX, pY, dX, dY, gFont, text, radioButton)
        return groupBox

    def myDataGridView(self, name, pX, pY, font, tableDefault):
        """Añadir una tabla editable"""
        datagrid = MyDataGridView(self, name, pX, pY, font, tableDefault)
        self.dataGridView1 = datagrid
        return self.dataGridView1

    def botonPulsado(self, sender, args):
        """Evento al pulsar el boton"""

        if sender.Name == 'continuar':
            for control in self.Controls:
                if type(control) == wforms.TextBox:
                    self.controlInfo['TextBox'][control.Name] = control.Text
                    # self.controlInfo['TextBox'].Add(control.Name, self.controlInfo['TextBox'][control.Text])
                elif type(control) == wforms.ComboBox:
                    self.controlInfo['ComboBox'][control.Name] = control.SelectedItem
                elif type(control) == wforms.CheckBox:
                    self.controlInfo['CheckBox'][control.Name] = control.Checked
                elif type(control) == ListView:
                    self.controlInfo['ListView'][control.Name] = [i.Text for i in
                                                                  control.CheckedItems]
                elif type(control) == DataGridView:
                    self.controlInfo['DataGridView'] = [x.ItemArray for x in
                                                        control.DataSource.Rows]
                elif type(control) == wforms.GroupBox:
                    for ctrl in control.Controls:
                        if type(ctrl) == wforms.CheckBox:
                            if ctrl.Name.split('-')[0] not in self.controlInfo['CheckBox']:
                                self.controlInfo['CheckBox'][ctrl.Name.split('-')[0]] = {}
                            self.controlInfo['CheckBox'][ctrl.Name.split('-')[0]][
                                ctrl.Name.split('-')[1]] = ctrl.Checked
                        if type(ctrl) == wforms.RadioButton:
                            if ctrl.Name.split('-')[0] not in self.controlInfo['RadioButton']:
                                self.controlInfo['RadioButton'][ctrl.Name.split('-')[0]] = {}
                            self.controlInfo['RadioButton'][ctrl.Name.split('-')[0]][
                                ctrl.Name.split('-')[1]] = ctrl.Checked
                        if type(ctrl) == wforms.TextBox:
                            if ctrl.Name.split('-')[0] not in self.controlInfo['TextBox']:
                                self.controlInfo['TextBox'][ctrl.Name.split('-')[0]] = {}
                            self.controlInfo['TextBox'][ctrl.Name.split('-')[0]][
                                ctrl.Name.split('-')[1]] = ctrl.Text
                else:
                    pass
            self.controlInfo['Button'] = sender.Name

            if 'CheckBox' in self.controlInfo:
                if 'Parametros a incluir en la descripción' in self.controlInfo['CheckBox']:
                    if len([v for v in self.controlInfo['CheckBox'][
                        'Parametros a incluir en la descripción'].values()
                            if v is True]) <= self.checkboxCount:
                        self.Close()
                    else:
                        MessageBox.Show("Por favor, seleccione máximo {} parámetros a incluir"
                                        .format(self.checkboxCount), " ", MessageBoxButtons.OK)
                else:
                    self.Close()

        elif sender.Name == 'cancelar':
            self.controlInfo['Cancel'] = sender.Name
            if MessageBox.Show(self, "Proceso cancelado", " ", MessageBoxButtons.OK,
                               MessageBoxIcon.Stop) == DialogResult.OK:
                self.Close()

        elif sender.Name == 'load':
            self.controlInfo['Load'] = sender.Name
            self.dataGridView1.load_table()

        elif sender.Name == 'save':
            self.controlInfo['Save'] = sender.Name
            self.dataGridView1.save_table()

        elif sender.Name == 'deleterows':
            self.controlInfo['Delete rows'] = sender.Name
            self.dataGridView1.delete_rows()


class MyLabel(wforms.Label):
    """Clase de etiqueta"""

    def __init__(self, ventana, pX, pY, font, text, dX=None, dY=None, borderstyle=None):
        """Constructor de la clase MyLabel"""

        self.label = wforms.Label()
        self.label.Location = Point(pX, pY)
        self.label.Font = font
        self.label.Text = text
        self.label.AutoSize = True
        self.label.BringToFront()

        if dX and dY:
            self.label.AutoSize = False
            self.label.ClientSize = Size(dX, dY)

        if borderstyle:
            self.label.BorderStyle = wforms.BorderStyle.Fixed3D

        ventana.Controls.Add(self.label)


class MyBorder(wforms.GroupBox):
    """Clase de etiqueta"""

    def __init__(self, ventana, pX, pY, dX, dY, font, text):
        """Constructor de MyGroupBox"""

        self.border = wforms.GroupBox()
        self.border.Location = Point(pX, pY)
        self.border.ClientSize = Size(dX, dY)
        self.border.Font = font
        self.border.Text = text

        ventana.Controls.Add(self.border)


class MyTextBoxMulty(wforms.TextBox):
    """Crear cuadros de texto"""

    def __init__(self, ventana, name, pX, pY, dX, dY, font, text='',
                 multiline=None, scrollBars=None):
        """Constructor de MyTextBox"""

        self.textBoxMulty = wforms.TextBox()
        self.textBoxMulty.Name = name
        self.textBoxMulty.Location = Point(pX, pY)
        self.textBoxMulty.ClientSize = Size(dX, dY)
        self.textBoxMulty.Font = font
        self.textBoxMulty.Text = text

        if multiline:
            self.textBoxMulty.Multiline = True
            if scrollBars:
                self.textBoxMulty.ScrollBars = wforms.ScrollBars.Vertical

        self.textBoxMulty.BorderStyle = wforms.BorderStyle.FixedSingle

        ventana.Controls.Add(self.textBoxMulty)


class MyComboBox(wforms.ComboBox):
    """Crea un Combo Box"""

    def __init__(self, ventana, name, pX, pY, font, text, elements,
                 dX=None, dY=None, flatStyle=None):
        """Constructor del objeto MyComboBox"""

        self.comboBox = wforms.ComboBox()
        self.comboBox.Name = name
        self.comboBox.Location = Point(pX, pY)
        self.comboBox.Font = font
        self.comboBox.Items.Insert(0, text)
        self.comboBox.SelectedIndex = 0
        self.comboBox.DropDownStyle = wforms.ComboBoxStyle.DropDownList

        if dX and dY:
            self.comboBox.AutoSize = False
            self.comboBox.ClientSize = Size(dX, dY)

        if flatStyle:
            self.comboBox.FlatStyle = flatStyle

        if type(elements) == list:
            listAux = elements
        else:
            listAux = [elements]

        [self.comboBox.Items.Add(element) for element in listAux]
        ventana.Controls.Add(self.comboBox)


class MyGroupBox(wforms.GroupBox):
    """Crea un GroupBox"""

    def __init__(self, ventana, pX, pY, dX, dY, font, text, elements):
        """Constructor de MyGroupBox"""

        self.groupBox = wforms.GroupBox()
        self.groupBox.Location = Point(pX, pY)
        self.groupBox.ClientSize = Size(dX, dY)
        self.groupBox.Font = font
        self.groupBox.Text = text

        [self.groupBox.Controls.Add(element) for element in elements]
        ventana.Controls.Add(self.groupBox)


class MyListView(ListView):
    """Crea un ListView"""

    def __init__(self, ventana, name, pX, pY, dX, dY, font, elements):
        """Constructor de MyListView"""

        self.listView = ListView()
        self.listView.Name = name
        self.listView.Location = Point(pX, pY)
        self.listView.ClientSize = Size(dX, dY)
        self.listView.Font = font
        self.listView.CheckBoxes = True
        self.listView.View = View.Details

        for encabezo, valores in elements.items():
            if encabezo.split('-')[0] == '0':
                self.listView.Columns.Add(encabezo.split('-')[-1], -2, HorizontalAlignment.Center)
                [self.listView.Items.Add(x) for x in valores]

        for encabezo, valores in elements.items():
            if encabezo.split('-')[0] == '0':
                continue
            else:
                self.listView.Columns.Insert.Overloads[Int32, String, Int32, HorizontalAlignment](
                    int(encabezo.split('-')[0]), encabezo.split('-')[1], -2,
                    HorizontalAlignment.Center)
                for index, atributo in enumerate(valores):
                    self.listView.Items[index].SubItems.Add(atributo)

        ventana.Controls.Add(self.listView)


class MyButton(Button):
    """Crea un objeto MyButton"""

    def __init__(self, ventana, name, pX, pY, font, text,
                 dX=None, dY=None, flatStyle=None):
        """Constructor de MyButton"""

        self.button = Button()
        self.button.Name = name
        self.button.Location = Point(pX, pY)
        self.button.Font = font
        self.button.Text = text
        self.button.AutoSize = True
        self.button.BackColor = Color.Gainsboro

        if dX and dY:
            self.button.AutoSize = False
            self.button.ClientSize = Size(dX, dY)

        if flatStyle:
            self.button.FlatStyle = wforms.FlatStyle.Flat
            self.button.UseVisualStyleBackColor = True
            self.button.FlatAppearance.BorderColor = Color.Gray
            self.button.FlatAppearance.MouseOverBackColor = Color.LightBlue

        self.button.Click += ventana.botonPulsado

        ventana.Controls.Add(self.button)


class MyMessageBox(Form):
    """Clase para crear interfaces graficas"""

    def __init__(self, message, title):
        self.Text = title
        self.label = wforms.Label()
        self.label.Text = message
        self.label.Location = Point(50, 50)
        self.label.Height = 30
        self.label.Width = 200
        self.label.AutoSize = True
        self.CenterToScreen()
        self.count = 0

        button = Button()
        button.Text = "OK"
        button.Location = Point(50, 100)

        button.Click += self.buttonPressed

        self.Controls.Add(self.label)
        self.Controls.Add(button)

    def buttonPressed(self, sender, args):
        self.Close()

    def run(self):
        """Muestra la interfaz en la ventana"""
        wforms.Application.Run(self)


class MyDataGridView(DataGridView):
    """Crea un GroupBox"""

    def __init__(self, ventana, name, pX, pY, font, tableDefault):
        self._dataGridView = DataGridView()
        # Add to main form
        ventana.Controls.Add(self._dataGridView)
        # Title
        self._dataGridView.Text = name
        # Set user options
        self._dataGridView.AllowUserToOrderColumns = True
        self._dataGridView.AllowUserToDeleteRows = True
        self._dataGridView.AllowUserToResizeRows = False
        self._dataGridView.MultiSelect = True
        # Appareance
        fontHeader = font
        if font.Style != FontStyle.Bold: fontHeader = myFont(font.Name, font.Size, FontStyle.Bold)
        self._dataGridView.ColumnHeadersDefaultCellStyle.Font = fontHeader
        self._dataGridView.DefaultCellStyle.Font = font
        self._dataGridView.TabIndex = 1
        self._dataGridView.ColumnHeadersVisible = True
        self._dataGridView.ColumnHeadersBorderStyle = wforms.DataGridViewHeaderBorderStyle.Raised
        self._dataGridView.CellBorderStyle = wforms.DataGridViewCellBorderStyle.Single
        self._dataGridView.SelectionMode = wforms.DataGridViewSelectionMode.FullRowSelect
        self._dataGridView.EditMode = wforms.DataGridViewEditMode.EditOnEnter
        # Color
        self._dataGridView.DefaultCellStyle.ForeColor = Color.Black
        self._dataGridView.DefaultCellStyle.BackColor = Color.White
        self._dataGridView.DefaultCellStyle.SelectionForeColor = Color.Black
        self._dataGridView.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue
        # Set data
        self._dataGridView.DataSource = tableDefault
        # Size
        self._dataGridView.Size = Size(810, 157)
        self.MinimumSize = Size(515, 200)
        self._dataGridView.Location = Point(pX, pY)
        self._dataGridView.AutoSizeColumnsMode = wforms.DataGridViewAutoSizeColumnsMode.AllCells
        self._dataGridView.Columns[0].AutoSizeMode = wforms.DataGridViewAutoSizeColumnsMode.Fill
        self._dataGridView.Update()

    def load_table(self):
        fileContent, filePath = "", ""
        openFileDialog = wforms.OpenFileDialog()
        openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        openFileDialog.FilterIndex = 2
        openFileDialog.RestoreDirectory = True
        if openFileDialog.ShowDialog() == DialogResult.OK:
            filePath = openFileDialog.FileName
            # header = [x.ColumnName for x in self._dataGridView.DataSource.Columns]
            fileContent = array_to_datatable(get_data_txt(filePath))
            self._dataGridView.DataSource = fileContent

        MessageBox.Show("Datos cargados desde: " + filePath, "Tabla cargada", MessageBoxButtons.OK)
        self._dataGridView.Update()
        return fileContent

    def save_table(self):
        fileContent, filePath = "", ""
        saveFileDialog = wforms.SaveFileDialog()
        saveFileDialog.OverwritePrompt = True
        saveFileDialog.RestoreDirectory = True
        saveFileDialog.FileName = "Clasificacion_Muros"
        saveFileDialog.DefaultExt = "txt"
        saveFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        textoinicial = "# Archivo de clasificacion de elementos\n" + "#" * 58
        fileContent = datatable_to_string(self._dataGridView.DataSource, textoinicial)
        if saveFileDialog.ShowDialog() == DialogResult.OK:
            filePath = saveFileDialog.FileName
            File.WriteAllLines(filePath, fileContent)
        MessageBox.Show("Datos guardados en: " + filePath, "Tabla guardada", MessageBoxButtons.OK)
        return filePath

    def delete_rows(self):
        totalrows = self._dataGridView.RowCount
        count = self._dataGridView.SelectedRows.Count
        if MessageBox.Show("¿Está seguro de que desea borrar las clasificaciones seleccionadas?",
                           "Aviso", MessageBoxButtons.YesNo) == DialogResult.Yes:
            if count > 0 and totalrows - count > 0:
                for row in self._dataGridView.SelectedRows:
                    self._dataGridView.DataSource.Rows.RemoveAt(row.Index)
            else:
                MessageBox.Show("No hay filas que borrar")

            self._dataGridView.AutoResizeColumns()
            self._dataGridView.Update()
        self._dataGridView.Update()
        return count
