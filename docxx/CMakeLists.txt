cmake_minimum_required(VERSION 3.22)
project(docxx)

set(CMAKE_CXX_STANDARD 14)

include_directories(thirdparty/pugixml thirdparty/zip)

add_library(docxx docxx.cpp utils.cpp thirdparty/pugixml/pugixml.cpp thirdparty/zip/zip.c)
