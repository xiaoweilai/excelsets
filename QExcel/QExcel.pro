#-------------------------------------------------
#
# Project created by QtCreator 2015-05-06T21:59:01
#
#-------------------------------------------------

#QT       += core gui

#TARGET = QExcel
#TEMPLATE = app


#SOURCES += main.cpp\
#        mainwindow.cpp \
#    qexcel.cpp

#HEADERS  += mainwindow.h \
#    qexcel.h

#FORMS    += mainwindow.ui





CONFIG += qaxcontainer

QT       += core

QT       -= gui

TARGET = QExcel
CONFIG   += console
CONFIG   -= app_bundle

TEMPLATE = app


SOURCES += main.cpp \
    qexcel.cpp

HEADERS += \
    qexcel.h
