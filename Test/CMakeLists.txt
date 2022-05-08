cmake_minimum_required(VERSION 3.22)
project(Test)

set(CMAKE_CXX_STANDARD 17)

include_directories(../docxx
                    ../docxx/thirdparty/pugixml
                    ../docxx/thirdparty/zip
)

link_directories("../docxx/cmake-build-${CMAKE_BUILD_TYPE}")

link_libraries(docxx)

add_executable(Test main.cpp)
