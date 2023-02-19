# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Info
                                 A QGIS plugin
 Make infomation about land
 Generated by Plugin Builder: http://g-sherman.github.io/Qgis-Plugin-Builder/
                              -------------------
        begin                : 2022-12-03
        git sha              : $Format:%H$
        copyright            : (C) 2022 by LevchenkoV
        email                : 180982@mail.ru
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication, Qt, QVariant
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction
from qgis.utils import iface
from qgis.core import QgsVectorLayer, QgsField
import datetime as dt
import os
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QFileDialog
# from PyQt5 import QtGui, QtCore
from fpdf import FPDF


# Initialize Qt resources from file resources.py
from .resources import *

# Import the code for the DockWidget
from .isogd_dockwidget import InfoDockWidget
import os.path


class Info:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface

        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)

        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'Info_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)
            QCoreApplication.installTranslator(self.translator)

        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&ISOGD')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'Info')
        self.toolbar.setObjectName(u'Info')

        #print "** INITIALIZING Info"

        self.pluginIsActive = False
        self.dockwidget = None


    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('Info', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):
        """Add a toolbar icon to the toolbar.

        :param icon_path: Path to the icon for this action. Can be a resource
            path (e.g. ':/plugins/foo/bar.png') or a normal file system path.
        :type icon_path: str

        :param text: Text that should be shown in menu items for this action.
        :type text: str

        :param callback: Function to be called when the action is triggered.
        :type callback: function

        :param enabled_flag: A flag indicating if the action should be enabled
            by default. Defaults to True.
        :type enabled_flag: bool

        :param add_to_menu: Flag indicating whether the action should also
            be added to the menu. Defaults to True.
        :type add_to_menu: bool

        :param add_to_toolbar: Flag indicating whether the action should also
            be added to the toolbar. Defaults to True.
        :type add_to_toolbar: bool

        :param status_tip: Optional text to show in a popup when mouse pointer
            hovers over the action.
        :type status_tip: str

        :param parent: Parent widget for the new action. Defaults None.
        :type parent: QWidget

        :param whats_this: Optional text to show in the status bar when the
            mouse pointer hovers over the action.

        :returns: The action that was created. Note that the action is also
            added to self.actions list.
        :rtype: QAction
        """

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action


    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        icon_path = ':/plugins/isogd/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u''),
            callback=self.run,
            parent=self.iface.mainWindow())

    #--------------------------------------------------------------------------

    def onClosePlugin(self):
        """Cleanup necessary items here when plugin dockwidget is closed"""

        #print "** CLOSING Info"

        # disconnects
        self.dockwidget.closingPlugin.disconnect(self.onClosePlugin)

        # remove this statement if dockwidget is to remain
        # for reuse if plugin is reopened
        # Commented next statement since it causes QGIS crashe
        # when closing the docked window:
        # self.dockwidget = None

        self.pluginIsActive = False


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""

        #print "** UNLOAD Info"

        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&ISOGD'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar

    #--------------------------------------------------------------------------

    def run(self):
        """Run method that loads and starts the plugin"""

        if not self.pluginIsActive:
            self.pluginIsActive = True

            print("** STARTING Info")

            # dockwidget may not exist if:
            #    first run of plugin
            #    removed on close (see self.onClosePlugin method)
            if self.dockwidget == None:
                # Create the dockwidget (after translation) and keep reference
                self.dockwidget = InfoDockWidget()

            # connect to provide cleanup on closing of dockwidget
            self.dockwidget.closingPlugin.connect(self.onClosePlugin)

            # show the dockwidget
            # TODO: fix to allow choice of dock location
            self.iface.addDockWidget(Qt.TopDockWidgetArea, self.dockwidget)
            self.dockwidget.show()

            # -----Проверяет исходный объект на пересечение с объектами слоя
            # -----Возвращает список полей слоя и список пересекаемых объектов
            class CheckLayers:
                def __init__(self, pathLayer, SelectedFeatureGeometry):
                    layer_check = QgsVectorLayer(pathLayer, pathLayer, "ogr")
                    fields = layer_check.fields()  # получаем список имен полей слоя
                    self.fields = fields
                    Features = layer_check.getFeatures()  # получаем все объекты слоя
                    k = 0
                    feat_intersect = []
                    procs = []
                    for feat in Features:  # перебираем объекты
                        geometry = feat.geometry()  # получаем геометрию объекта
                        rezult_intersects = SelectedFeatureGeometry.intersects(geometry)  # пересекаем объекты
                        rezult_intersection = SelectedFeatureGeometry.intersection(geometry)
                        if rezult_intersects and (rezult_intersection.area() > 0):
                            s = float(rezult_intersection.area())
                            s1 = float(SelectedFeatureGeometry.area())
                            proc = str(round(s * 100/s1, 5))
                            procs.append(proc)
                            feat_intersect.append(feat) # создаем список объектов, с которым есть пересечение
                            k += 1
                    if k == 0: print("      пересечений не выявлено")
                    self.feat_intersect = feat_intersect
                    self.procs = procs

            # --- Создает таблицу для отчета, если пересечений нет, то таблица заполняется None
            def make_table(fields, feat_intersect, procs):
                headers = []                        #------Заголовки таблицы
                for field in fields:
                    headers.append(field.name())
                headers.append("%пересеч")
                rows = []                          # -------Строки таблицы
                if feat_intersect != []: #----Строки таблицы, если есть пересечения
                    for (feat, proc) in zip(feat_intersect, procs):
                        row = []
                        for field in fields:
                                row.append(str(feat.attribute(field.name())))
                        row.append(proc)
                        rows.append(row)
                else:  #----Строки таблицы, если нет пересечений, заполняется None
                    row = []
                    for header in headers:
                        header = "None"
                        row.append(header)
                    rows.append(row)

                return headers, rows

            class PDF(FPDF):  # ----печать данных без вывода в файл

                def footer(self):
                    # Go to 1.5 cm from bottom
                    self.set_y(-20)
                    # Select Arial italic 8
                    self.set_text_color(0)
                    self.set_font('Sans', 'I', 8)
                    # Print centered page number
                    self.cell(0, 10, 'Page %s' % self.page_no(), 0, 0, 'C')

                def head(self, Layer_Source_Name):
                    self.cell(200, 6, Layer_Source_Name, 1, 0, "C")
                    self.ln(40)

                def basic_table(self, headings, rows, col_width, pathLayer):

                    # цвет текста
                    self.set_text_color(255)
                    # цвет линий таблицы
                    self.set_draw_color(224, 235, 255)
                    # ширина линии
                    self.set_line_width(0.3)
                    # цвет заливки
                    self.set_fill_color(0, 51, 102)

                    # если пересечений нет - убрать цвет
                    if rows[0][0] == 'None':
                        self.set_fill_color(125, 125, 125)

                    # печать названия слоя по которому идет проверка
                    self.cell(sum(col_width), 6, pathLayer, 1, 0, "C", True)
                    self.ln(6)

                    # печать заголовка таблицы
                    self.set_fill_color(255, 102, 0)

                    # если пересечений нет - убрать цвет
                    if rows[0][0] == 'None':
                        self.set_fill_color(190, 190, 190)

                    for heading, width1 in zip(headings, col_width):
                        self.cell(width1, 6, heading, 1, 0, "C", True)
                    self.ln()

                    self.set_fill_color(224, 235, 255)
                    self.set_text_color(0)

                    fill = False
                    for row in rows:
                        for col, width2 in zip(row, col_width):
                            a = int(200*width2/266) - 2
                            col_new = col.replace('PyQt5.QtCore.QDate', '') #---убираем PyQt5.QtCore.QDate из вывода даты
                            col_new2 = col_new[:a]                          #---обрезаем длину строки по ширине колонки
                            self.cell(width2, 6, col_new2, "LR", 0, "L", fill)
                        self.ln()
                        fill = not fill
                    self.cell(sum(col_width), 0, "", "T")
                    self.ln(6)


            def window_error():
                error = QMessageBox()
                error.setWindowTitle("Ошибка")
                error.setText("Не выбран объект для проверки!!! Выберите полигон")
                error.setStandardButtons(QMessageBox.Ok)
                error.exec_()

            def get_dir_name():
                name_dir = QFileDialog.getExistingDirectory(None, 'Выберите папку для сохранения Справки')
                return name_dir

            def isogd():
                data_time = dt.datetime.now()
                time_now = data_time.strftime("%H%M%S")
                print(time_now)
                pathLayers = [
                              # 'D:/Эксперименты в Qgis/Spravka/Охранная зона 1.shp',
                              'T:/Номенклатура/Лист_500.TAB'
                              # 'T:/NOVOKUZ/Красные_линии_полигон.TAB',
                              #
                              # 'T:/Территориальные зоны/Терр_зоны_на_ГКУМСК42зона2.TAB',
                              # 'O:/_Str-ISOGD/ПЗиЗ/9. ПЗиЗ Редакция октябрь 2022/Вектор/Общий объединенный ноябрь 2022 Вектор.TAB',
                              # 'T:/NOVOKUZ/ФГУ участки/KK.TAB',
                              #
                              # 'T:/NOVOKUZ/Участки.TAB',
                              # 'T:/NOVOKUZ/ФГУ участки/ACTUAL_LAND.TAB',
                              #
                              # 'T:/NOVOKUZ/ФГУ участки/ACTUAL_OKSN.TAB',
                              # 'T:/ADRES/Строения.TAB',
                              #
                              # 'T:/NOVOKUZ/Схема расположения.TAB',
                              # 'T:/NOVOKUZ/Предварительно согласованные.TAB',
                              # 'T:/NOVOKUZ/Проекты планировок и межеваний.TAB',
                              #
                              # 'T:/NOVOKUZ/ФГУ участки/ACTUAL_ZOUIT.TAB',
                              # 'T:/NOVOKUZ/ФГУ участки/ЗОУИТ.TAB',
                              #
                              # 'T:/NOVOKUZ/Водоохранные зоны Томь, Аба, Горбуниха, Бунгур, Кондома (ЦИТ Барнаул)/Береговая_Прибрежная_Водоохранная.TAB',
                              # 'T:/NOVOKUZ/Приаэродромная территория МСК42зона2.tab',
                              # 'T:/NOVOKUZ/Санзоны новые.TAB',
                              # 'T:/NOVOKUZ/Охранная зона.TAB',
                              #
                              # 'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Границы территорий объектов Культурного наследия.TAB',
                              # 'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Зоны охраны объектов культурного наследия.TAB',
                              # 'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Объекты культурного наследия на ГКУ.TAB',
                              # 'T:/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Объекты культурного наследия.TAB',
                              #
                              # 'T:/NOVOKUZ/Разрешение на использование ЗУ.TAB',
                              # 'T:/NOVOKUZ/Схема НТО.TAB',
                              # 'T:/NOVOKUZ/Схема НТО/Договоры НТО.TAB',
                              # 'T:/NOVOKUZ/Сервитуты.TAB',
                              # 'T:/NOVOKUZ/Ветхое и аварийное жильё.TAB',
                              #
                              # 'T:/NOVOKUZ/Запреты/Арест судебные приставы/Запрет судебных приставов.TAB',
                              # 'T:/NOVOKUZ/Запреты/Реконструкция южного въезда/Южная автомагистраль.TAB',
                              # 'T:/NOVOKUZ/Запреты/Строительство Байдаевского моста/Байдаевский мост изъятие.TAB'
                              # 'D:/Эксперименты в Qgis/Spravka/Охранная зона 2.shp'
                                ]
                col_widths = [
                              # [20, 20, 14],
                              [43, 43, 14]  #T:/Номенклатура /Лист_500.TAB
                              # [10, 126, 14],  # T:/NOVOKUZ/Красные_линии_полигон.TAB
                              #
                              # [46, 46, 47, 47, 14],  # T:/Территориальные зоны/Терр_зоны_на_ГКУМСК42зона2.TAB
                              # [93, 93, 14], # O:/_Str-ISOGD/ПЗиЗ/9. ПЗиЗ Редакция октябрь 2022/Вектор/Общий объединенный ноябрь 2022 Вектор.TAB
                              # [20, 20, 20, 146, 60, 20, 20, 20, 20, 20, 20, 14],  # T:/NOVOKUZ/ФГУ участки/KK.TAB
                              #
                              # [18, 30, 80, 50, 45, 20, 32, 22, 22, 40, 10, 17, 14], # T:/NOVOKUZ/Участки.TAB
                              # [27, 25, 190, 50, 18, 21, 36, 19, 14],  # T:/NOVOKUZ/ФГУ участки/ACTUAL_LAND.TAB
                              #
                              # [38, 38, 38, 38, 38, 38, 38, 40, 40, 40, 14],  # T:/NOVOKUZ/ФГУ участки/ACTUAL_OKSN.TAB
                              # [64, 66, 64, 64, 64, 64, 14],  # T:/ADRES/Строения.TAB
                              #
                              # [10, 50, 10, 10, 60, 60, 10, 10, 10, 10, 10, 10, 20, 10, 10, 10, 46, 10, 10, 10, 14],  # T:/NOVOKUZ/Схема расположения.TAB
                              # [42, 72, 64, 42, 20, 20, 42, 42, 42, 14],  # T:/NOVOKUZ/Предварительно согласованные.TAB
                              # [30, 70, 26, 28, 28, 26, 26, 26, 26, 100, 14],  # T:/NOVOKUZ/Проекты планировок и межеваний.TAB
                              #
                              # [100, 36, 40, 150, 40, 20, 14],  # T:/NOVOKUZ/ФГУ участки/ACTUAL_ZOUIT.TAB
                              # [10, 25, 150, 156, 23, 22, 14],  # T:/NOVOKUZ/ФГУ участки/ЗОУИТ.TAB
                              #
                              # [20, 266, 14], # T:/NOVOKUZ/Водоохранные зоны Томь, Аба, Горбуниха, Бунгур, Кондома (ЦИТ Барнаул)/Береговая_Прибрежная_Водоохранная.TAB
                              # [178, 108, 14], #   T:/NOVOKUZ/Приаэродромная территория МСК42зона2.tab
                              # [178, 108, 14],   # T:/NOVOKUZ/Санзоны новые.TAB
                              # [20, 75, 75, 75, 15, 26, 14],  # T:/NOVOKUZ/Охранная зона.TAB
                              #
                              # [100, 186, 100, 14], #E:/26.11.2022/T_26.11.2022/NOVOKUZ/ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ/Границы территорий объектов Культурного наследия.TAB
                              # [40, 50, 46, 250, 14], #E:\26.11.2022\T_26.11.2022\NOVOKUZ\ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ\Зоны охраны объектов культурного наследия.TAB
                              # [38, 38, 38, 38, 39, 39, 39, 39, 39, 39, 14], # 'E:\26.11.2022\T_26.11.2022\NOVOKUZ\ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ\Объекты культурного наследия на ГКУ.TAB'
                              # [40, 40, 60, 20, 40, 40, 16, 110, 10, 10, 14], #E:\26.11.2022\T_26.11.2022\NOVOKUZ\ЗОНЫ КУЛЬТУРНОГО НАСЛЕДИЯ\Объекты культурного наследия.TAB
                              #
                              # [10, 60, 25, 20, 15, 22, 26, 70, 70, 19, 27, 22, 14], # T:/NOVOKUZ/Разрешение на использование ЗУ.TAB
                              # [20, 50, 15, 15, 20, 20, 70, 30, 66, 20, 20, 20, 20, 14], # T:/NOVOKUZ/Схема НТО.TAB'
                              # [20, 66, 20, 20, 20, 20, 20, 20, 20, 70, 45, 25, 20, 14], # T:/NOVOKUZ/Схема НТО/Договоры НТО.TAB'
                              # [16, 25, 20, 20, 18, 12, 25, 25, 50, 20, 15, 10, 40, 50, 20, 20, 14],  # T:/NOVOKUZ/Сервитуты.TAB
                              # [56, 55, 55, 55, 55, 55, 55, 14], #T:/NOVOKUZ/Ветхое и аварийное жильё.TAB
                              #
                              # [66, 64, 64, 64, 64, 64, 14], # T:/NOVOKUZ/Запреты/Арест судебные приставы/Запрет судебных приставов.TAB
                              # [18, 30, 80, 50, 45, 20, 32, 22, 22, 40, 10, 17, 14], # T:/NOVOKUZ/Запреты/Реконструкция южного въезда/Южная автомагистраль.TAB
                              # [20, 20, 46, 40, 40, 20, 20, 20, 20, 70, 25, 25, 20, 14], # T:/NOVOKUZ/Запреты/Строительство Байдаевского моста/Байдаевский мост изъятие.TAB
                              # [20, 20, 14]
                            ]
                Layer = iface.activeLayer()
                SelectedFeatures = Layer.getSelectedFeatures()
                Layer_Source_Name = Layer.sourceName()
                m = 0
                for SelectedFeature in SelectedFeatures:  # перебираем объекты, выводим ошибку, если не выбрано ничего
                    m += 1
                if m == 0:
                    window_error()
                else:
                    SelectedFeatures = Layer.getSelectedFeatures()
                    for SelectedFeature in SelectedFeatures:  # перебираем объекты
                        # ----формирование имени файла из семантики объекта слоя
                        fields_source = SelectedFeature.fields()
                        field_0 = str(SelectedFeature.attribute(fields_source[0].name()))  # получаем семантику поля [0]
                        field_4 = str(SelectedFeature.attribute(fields_source[4].name()))  # получаем семантику поля [1]
                        dir_name = get_dir_name()

                        file_name = dir_name + '/Вх№' + \
                                    field_0 + ' ' + field_4 + ' ' + str(time_now) + ".pdf"

                        SelectedFeatureGeometry = SelectedFeature.geometry()  # получаем геометрию объекта


                        # --------------------подготовка к печати pdf
                        pdf = PDF(orientation="L", unit="mm", format="A3")

                        # включаем TTF шрифты, поддерживающие кириллицу
                        pdf.add_font("Sans", style="",
                                     fname="C:/Fonts/NotoSans-Regular.ttf",
                                     uni=True)
                        pdf.add_font("Sans", style="B",
                                     fname="C:/Fonts/NotoSans-Bold.ttf",
                                     uni=True)
                        pdf.add_font("Sans", style="I",
                                     fname="C:/Fonts/NotoSans-Italic.ttf",
                                     uni=True)
                        pdf.add_font("Sans", style="BI",
                                     fname="C:/Fonts/NotoSans-BoldItalic.ttf",
                                     uni=True)
                        # настройка шрифта
                        pdf.set_font("Sans", size=7)
                        pdf.add_page()

                        pdf.head(Layer_Source_Name)

                        progress = QProgressDialog("ИСОГД работает изо всех сил...", "Cancel", 0, len(pathLayers)) # процесс выполнения
                        progress.setWindowModality(Qt.WindowModal)

                        i = 0
                        for pathLayer, col_width in zip(pathLayers, col_widths):
                            a = CheckLayers(pathLayer, SelectedFeatureGeometry)
                            col_names, data = make_table(a.fields, a.feat_intersect, a.procs)
                            pdf.basic_table(col_names, data, col_width, pathLayer)
                            progress.setValue(i)
                            i += 1
                        progress.setValue(len(pathLayers))

                        #---- вывод напечатанного в файл и его открытие
                        pdf.output(file_name)
                        os.startfile(file_name)

            self.dockwidget.pushButton.clicked.connect(isogd)