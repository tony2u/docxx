#include <iostream>
#include "docxx.hpp"
#ifdef _MSC_VER
#include <vld.h>
#endif

using namespace std;
using namespace docxx;

void test()
{
    Document doc("file.docx", true);
    auto lts = doc.tables();
    std::cout << "tables.size()=" << lts.size() << std::endl;
    auto tbl = lts.begin();
    std::cout << "tables[0].children().size()=" << tbl->children().size() << std::endl;
    auto lps = doc.paragraphs();
    std::cout << "paragraphs.size()=" << lps.size() << std::endl;
    auto para = lps.begin();
    std::cout << "paragraphs[0].children().size()=" << para->children().size() << std::endl;


}

struct UserInfo
{
    std::string imgSrc;    //��Ƭurl
    std::string imgWidth;  //���
    std::string imgHeight; //�߶�
    std::string name;      //����
    std::string ident;     //ִҵ֤��
    std::string certNo;    //�ʸ�֤��
    std::string unit;      //ִҵ����
    std::string sexy;      //�Ա�
    std::string age;       //����
    std::string group;     //����
    std::string email;     //��������
    std::string leader;    //����˾����
};
bool append_range_docx(const UserInfo& detail, bool convUTF8)
{
    Document doc("file.docx", true);

    //��ȡ��ǰword�еı��
    auto tbl = doc.tables().back();
    auto tr = tbl.rows().back().append_row(false); //�����һ��֮�� ׷��һ������
    auto c1 = tr.cells()[0];
    auto n = std::to_string(tbl.rows().size());
    c1.paragraphs().front().runs().front().add_text(n, convUTF8);
    auto c2 = tr.cells()[1];
    c2.paragraphs().front().runs().front().add_text(detail.name, convUTF8);
    auto c3 = tr.cells()[2];
    c3.paragraphs().front().runs().front().add_text(detail.email, convUTF8);
    auto c4 = tr.cells()[3];
    c4.paragraphs().front().runs().front().add_text(detail.unit, convUTF8);
    auto c5 = tr.cells()[4];
    c5.paragraphs().front().runs().front().add_text("", convUTF8);
    auto c6 = tr.cells()[5];
    c6.paragraphs().front().runs().front().add_text(detail.ident, convUTF8);
    char *stop_string = nullptr;
    auto wd = strtol(detail.imgWidth.c_str(), &stop_string, 10);
    auto ht = strtol(detail.imgHeight.c_str(), &stop_string, 10);
    auto c7 = tr.cells()[6];
    c7.paragraphs().front().runs().front().add_picture(doc, detail.imgSrc, wd, ht);

    cout << "����..." << endl;
    return doc.save();
}
void test2()
{
    UserInfo detail;
    detail.name = "aaa";
    detail.email = "aaa@abc.com";
    detail.unit = "�Ϻ���˾";
    detail.ident = "1234567890";
    detail.imgWidth = "110";
    detail.imgHeight = "140";
    detail.imgSrc = R"(E:\Temp\lawyer_cpp\Test\view.png)";
    append_range_docx(detail, true);
}

int main()
{
#ifdef _MSC_VER
    VLDSetOptions(VLD_OPT_TRACE_INTERNAL_FRAMES | VLD_OPT_SKIP_CRTSTARTUP_LEAKS, 256, 64);
#endif
    //test();
    test2();

#ifdef _MSC_VER
    //_CrtDumpMemoryLeaks();
    int leaks = VLDGetLeaksCount();
    VLDReportLeaks(); // at this point should report 9 leaks;
    return leaks;
#else
    return 0;
#endif
}
