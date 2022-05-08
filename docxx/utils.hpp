#ifndef DOCXX_UTILS_H
#define DOCXX_UTILS_H

#include <iostream>
#include <iostream>
#include <vector>
#include <sstream>
#include <cstdarg>
#include "pugixml.hpp"

unsigned int dice();
std::vector<std::string> split(const std::string& s, char delim);

//std::string &format(const char* format, ...);
//std::string string_format(const char* format, ...);

//char* UTF8ToGB2312(const char* utf8);
//char* GB2312ToUTF8(const char* gb2312);
void UTF8ToGB2312(const char* utf8, std::string& gb2312_str);
void GB2312ToUTF8(const char* gb2312, std::string& utf8_str);

// Hack on pugixml
// We need to write xml to std string (or char *)
// So overload the write function
struct xml_string_writer : pugi::xml_writer {
    std::string result;

    virtual void write(const void* data, size_t size) {
        result.append(static_cast<const char*>(data), size);
    }
};

#endif // DOCXX_UTILS_H
