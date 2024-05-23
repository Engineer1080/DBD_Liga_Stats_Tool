from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg, NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from PyQt5.QtWidgets import *
from PyQt5.Qt import QApplication, QUrl, QDesktopServices

import sys
import matplotlib
import openpyxl

matplotlib.use('Qt5Agg')


def skillgap(Team1_Points, Team2_Points):
    skillgaplist = list()
    for item1, item2 in zip(Team1_Points, Team2_Points):
        if item1 > item2:
            item = item1 - item2
        else:
            item = item2 - item1
        skillgaplist.append(item)
    return skillgaplist


class MplCanvas(FigureCanvasQTAgg):

    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super(MplCanvas, self).__init__(fig)


class MainWindow(QMainWindow):

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("DBD Tournament Statistics Tool")
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('File')
        editMenu = mainMenu.addMenu('Edit')
        viewMenu = mainMenu.addMenu('View')
        searchMenu = mainMenu.addMenu('Search')
        toolsMenu = mainMenu.addMenu('Tools')
        helpMenu = mainMenu.addMenu('Help')
        openButton = QAction(QIcon('exit24.png'), 'Open', self)
        openButton.setShortcut('Ctrl+O')
        openButton.setStatusTip('Open new .xlsx file')
        openButton.triggered.connect(self.dialog)
        fileMenu.addAction(openButton)
        exitButton = QAction(QIcon('exit24.png'), 'Exit', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip('Exit application')
        exitButton.triggered.connect(self.close)
        fileMenu.addAction(exitButton)
        help_action = QAction(QIcon("bug.png"), "Liga Wiki", self)
        help_action.triggered.connect(self.Help)
        helpMenu.addAction(help_action)
        self.show()

    def Help(self):
        url = QUrl("https://example.org")
        QDesktopServices.openUrl(url)

    def dialog(self):
        name = QFileDialog.getOpenFileName(self, 'Open File', filter="XLSX Files (*.xlsx)")
        plt = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        plt2 = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        wb = openpyxl.load_workbook(name[0], data_only=True)
        layout = QtWidgets.QVBoxLayout()
        sheet = wb.active
        killers = []
        killers_filtered = []
        killer_occurences = [0,0,0,0,0,0]
        killers.append(sheet['D17'].value)
        killers.append(sheet['I17'].value)
        killers.append(sheet['D33'].value)
        killers.append(sheet['I33'].value)
        killers.append(sheet['D49'].value)
        killers.append(sheet['I49'].value)
        for item in killers:
            if item not in killers_filtered:
                killers_filtered.append(item)
            killer_occurences[killers_filtered.index(item)] = killer_occurences[killers_filtered.index(item)] + 1
        killer_length = len(killers_filtered)
        difference = 6 - killer_length
        for x in range(difference):
            killer_occurences.pop()
        killerpie = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        killerpie.axes.pie(killer_occurences, labels=killers_filtered)
        killerpie.axes.set_xlabel("Killer Verteilung")
        maps = []
        maps_filtered = []
        maps_occurences = [0, 0, 0]
        maps.append(sheet['N57'].value)
        maps.append(sheet['N58'].value)
        maps.append(sheet['N59'].value)
        for item in maps:
            if item not in maps_filtered:
                maps_filtered.append(item)
            maps_occurences[maps_filtered.index(item)] = maps_occurences[maps_filtered.index(item)] + 1
        maps_length = len(maps_filtered)
        maps_difference = 3 - maps_length
        for x in range(maps_difference):
            maps_occurences.pop()
        mappie = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        mappie.axes.pie(maps_occurences, labels=maps_filtered)
        mappie.axes.set_xlabel("Map Verteilung")
        Team1_Name = sheet['B3'].value
        Team2_Name = sheet['G3'].value
        Game1_Match1_Round1_Team1_Survs = sheet['B15'].value
        Game1_Match1_Round1_Team2_Killer = sheet['D15'].value
        Game1_Match1_Round1_Team2_Killer_Name = sheet['D16'].value
        Game1_Match1_Round2_Team2_Survs = sheet['G15'].value
        Game1_Match1_Round2_Team1_Killer = sheet['I15'].value
        Game1_Match1_Round2_Team1_Killer_Name = sheet['I16'].value
        Game1_Match2_Round1_Team2_Survs = sheet['B31'].value
        Game1_Match2_Round1_Team1_Killer = sheet['D31'].value
        Game1_Match2_Round1_Team1_Killer_Name = sheet['D32'].value
        Game1_Match2_Round2_Team1_Survs = sheet['G31'].value
        Game1_Match2_Round2_Team2_Killer = sheet['I31'].value
        Game1_Match2_Round2_Team2_Killer_Name = sheet['I32'].value
        Game1_Match3_Round1_Team1_Survs = sheet['B47'].value
        Game1_Match3_Round1_Team2_Killer = sheet['D47'].value
        Game1_Match3_Round1_Team2_Killer_Name = sheet['D48'].value
        Game1_Match3_Round2_Team2_Survs = sheet['G47'].value
        Game1_Match3_Round2_Team1_Killer = sheet['I47'].value
        Game1_Match3_Round2_Team1_Killer_Name = sheet['I48'].value
        Team1_Survs_Points = []
        Team2_Survs_Points = []
        Team1_Survs_Points.append(Game1_Match1_Round1_Team1_Survs)
        Team1_Survs_Points.append(Game1_Match2_Round2_Team1_Survs)
        Team1_Survs_Points.append(Game1_Match3_Round1_Team1_Survs)
        Team2_Survs_Points.append(Game1_Match1_Round2_Team2_Survs)
        Team2_Survs_Points.append(Game1_Match2_Round1_Team2_Survs)
        Team2_Survs_Points.append(Game1_Match3_Round2_Team2_Survs)
        skillgapsurv = skillgap(Team1_Survs_Points, Team2_Survs_Points)
        rounds = [1, 2, 3]
        plt.axes.plot(rounds, Team1_Survs_Points, color='blue', linestyle='dashed', linewidth=3,
                      marker='o', markerfacecolor='blue', markersize=12, label=Team1_Name)
        plt.axes.plot(rounds, Team2_Survs_Points, color='green', linestyle='dashed', linewidth=3,
                      marker='o', markerfacecolor='green', markersize=12, label=Team2_Name)
        plt.axes.plot(rounds, skillgapsurv, color='grey', linestyle='dashed', linewidth=3,
                      marker='o', markerfacecolor='grey', markersize=12, label="Skill-Unterschied")
        plt.axes.set_xlabel("Runden")
        plt.axes.set_ylabel("Punktzahl")
        plt.axes.legend()
        Team1_Killer_Points = []
        Team2_Killer_Points = []
        Team1_Killer_Points.append(Game1_Match1_Round2_Team1_Killer)
        Team1_Killer_Points.append(Game1_Match2_Round1_Team1_Killer)
        Team1_Killer_Points.append(Game1_Match3_Round2_Team1_Killer)
        Team2_Killer_Points.append(Game1_Match1_Round1_Team2_Killer)
        Team2_Killer_Points.append(Game1_Match2_Round2_Team2_Killer)
        Team2_Killer_Points.append(Game1_Match3_Round1_Team2_Killer)
        skillgapkiller = skillgap(Team1_Killer_Points, Team2_Killer_Points)
        rounds = [1, 2, 3]
        plt2.axes.plot(rounds, Team1_Killer_Points, color='red', linestyle='dashed', linewidth=3,
                       marker='o', markerfacecolor='red', markersize=12, label=Team1_Name)
        plt2.axes.plot(rounds, Team2_Killer_Points, color='purple', linestyle='dashed', linewidth=3,
                       marker='o', markerfacecolor='purple', markersize=12, label=Team2_Name)
        plt2.axes.plot(rounds, skillgapkiller, color='grey', linestyle='dashed', linewidth=3,
                       marker='o', markerfacecolor='grey', markersize=12, label="Skill-Unterschied")
        plt2.axes.set_xlabel("Runden")
        plt2.axes.set_ylabel("Punktzahl")
        plt2.axes.legend()

        killerplayerspie = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        killer_players = []
        killer_players_filtered = []
        killer_players_occurences = [0, 0, 0, 0, 0, 0]
        killer_players.append(sheet['D16'].value)
        killer_players.append(sheet['I16'].value)
        killer_players.append(sheet['D32'].value)
        killer_players.append(sheet['I32'].value)
        killer_players.append(sheet['D48'].value)
        killer_players.append(sheet['I48'].value)
        for item in killer_players:
            if item not in killer_players_filtered:
                killer_players_filtered.append(item)
            killer_players_occurences[killer_players_filtered.index(item)] = killer_players_occurences[killer_players_filtered.index(item)] + 1
        killer_players_length = len(killer_players_filtered)
        difference = 6 - killer_players_length
        for x in range(difference):
            killer_players_occurences.pop()
        print(killer_players_occurences)
        print(killer_players)
        print(killer_players_filtered)
        killerplayerspie.axes.pie(killer_players_occurences, labels=killer_players_filtered)
        killerplayerspie.axes.set_xlabel("Killer-Spieler Verteilung")


        killer_average = MplCanvas(MainWindow, width=5, height=4, dpi=100)
        killer_points = []
        killer_points_filtered = []
        killer_points.append(sheet['D15'].value)
        killer_points.append(sheet['I15'].value)
        killer_points.append(sheet['D31'].value)
        killer_points.append(sheet['I31'].value)
        killer_points.append(sheet['D47'].value)
        killer_points.append(sheet['I47'].value)
        indexlist = []
        for killer_players_filtered in killer_players:
            indexlist.append(killer_players.index(killer_players_filtered))
            for v in indexlist:
                killer_points_filtered.append(killer_points[indexlist[v]])
        print(killer_points)
        print(indexlist)
        print(killer_points_filtered)

        New_Colors = ['pink','blue']
        killer_average.axes.bar(killer_players_filtered, killer_points_filtered, color=New_Colors)
        killer_average.axes.set_xlabel('Killer', fontsize=10)
        killer_average.axes.set_ylabel('Punkte-Durchschnitt', fontsize=10)

        toolbar = NavigationToolbar(plt, self)
        #toolbar2 = NavigationToolbar(plt2, self)
        layout.addWidget(toolbar)
        #layout.addWidget(toolbar2)
        layout.addWidget(plt)
        layout.addWidget(plt2)
        layout.addWidget(killerpie)
        layout.addWidget(mappie)
        layout.addWidget(killerplayerspie)
        layout.addWidget(killer_average)
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)
        self.show()


app = QtWidgets.QApplication(sys.argv)
w = MainWindow()
app.exec_()
