#include <iostream>
#include <iostream>
#include <vector>
#include <sstream>
#include <cstdarg>
#include <random>
#include <functional>

#include "pugixml.hpp"
#include "utils.hpp"
#include <windows.h>

unsigned int dice() {
    std::default_random_engine e(static_cast<unsigned int>(time(nullptr)));
    std::uniform_int_distribution<> u(1, 1000000);
    //auto diz = std::bind(u, e);
    //auto diz = [u, e] { return u(e); };
    //return diz();
    return u(e);
}

static void split(const std::string& s, char delim, std::vector<std::string>& elems) {
    std::stringstream ss(s);
    std::string item;
    while (std::getline(ss, item, delim)) {
        elems.push_back(item);
    }
}

std::vector<std::string> split(const std::string& s, char delim) {
    std::vector<std::string> elems;
    split(s, delim, elems);
    return elems;
}

//std::string &format(const char* format, ...)
//{
//    size_t size = 4096;
//    std::string buffer(size, '\0');
//    char* buffer_p = const_cast<char*>(buffer.data());
//    int expected = 0;
//    va_list ap;
//
//    while (true)
//    {
//        va_start(ap, format);
//        expected = vsnprintf(buffer_p, size, format, ap);
//
//        va_end(ap);
//        if (expected > -1 && expected <= static_cast<int>(size))
//        {
//            break;
//        }
//        else
//        {
//            /* Else try again with more space. */
//            if (expected > -1)    /* glibc 2.1 */
//                size = static_cast<size_t>(expected + 1); /* precisely what is needed */
//            else           /* glibc 2.0 */
//                size *= 2;  /* twice the old size */
//
//            buffer.resize(size);
//            buffer_p = const_cast<char*>(buffer.data());
//        }
//    }
//
//    // expected不包含字符串结尾符号，其值等于：strlen(buffer_p)
//    return std::string(buffer_p, expected > 0 ? expected : 0);
//}
//std::string string_format(const char* format, ...)
//{
//    va_list args;
//    va_start(args, format);
//    int count = 0;
//    count = vsnprintf(nullptr, 0, format, args);
//    va_end(args);
//
//    va_start(args, format);
//    char* buff = nullptr;
//    buff = (char*)malloc((count + 1) * sizeof(char));
//    vsnprintf(buff, (count + 1), format, args);
//    va_end(args);
//
//    std::string str(buff, count);
//    free(buff);
//    return str;
//}

//UTF-8到GB2312的转换
//char* UTF8ToGB2312(const char* utf8)
//{
//    int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
//    wchar_t* wstr = new wchar_t[len + 1];
//    memset(wstr, 0, len + 1);
//    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wstr, len);
//    len = WideCharToMultiByte(CP_ACP, 0, wstr, -1, NULL, 0, NULL, NULL);
//    char* str = new char[len + 1];
//    memset(str, 0, len + 1);
//    WideCharToMultiByte(CP_ACP, 0, wstr, -1, str, len, NULL, NULL);
//    if (wstr) delete[] wstr;
//    return str;
//}
void UTF8ToGB2312(const char* utf8, std::string& gb2312_str)
{
    int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    auto* wstr = new wchar_t[len + 1];
    memset(wstr, 0, len + 1);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wstr, len);
    len = WideCharToMultiByte(CP_ACP, 0, wstr, -1, nullptr, 0, nullptr, nullptr);
    char* str = new char[len + 1];
    memset(str, 0, len + 1);
    WideCharToMultiByte(CP_ACP, 0, wstr, -1, str, len, nullptr, nullptr);
    delete[] wstr;
    gb2312_str = str;
    delete[] str;
}

//GB2312到UTF-8的转换
//char* GB2312ToUTF8(const char* gb2312)
//{
//    int len = MultiByteToWideChar(CP_ACP, 0, gb2312, -1, NULL, 0);
//    wchar_t* wstr = new wchar_t[len + 1];
//    memset(wstr, 0, len + 1);
//    MultiByteToWideChar(CP_ACP, 0, gb2312, -1, wstr, len);
//    len = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, NULL, 0, NULL, NULL);
//    char* str = new char[len + 1];
//    memset(str, 0, len + 1);
//    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, str, len, NULL, NULL);
//    if (wstr) delete[] wstr;
//    return str;
//}
void GB2312ToUTF8(const char* gb2312, std::string& utf8_str)
{
    int len = MultiByteToWideChar(CP_ACP, 0, gb2312, -1, nullptr, 0);
    auto* wstr = new wchar_t[len + 1];
    memset(wstr, 0, len + 1);
    MultiByteToWideChar(CP_ACP, 0, gb2312, -1, wstr, len);
    len = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, nullptr, 0, nullptr, nullptr);
    char* str = new char[len + 1];
    memset(str, 0, len + 1);
    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, str, len, nullptr, nullptr);
    delete[] wstr;
    utf8_str = str;
    delete[] str;
}