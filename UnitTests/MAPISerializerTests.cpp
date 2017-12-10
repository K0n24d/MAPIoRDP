#include "stdafx.h"
#include "CppUnitTest.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

#include <boost\iostreams\device\back_inserter.hpp>
#include <boost\iostreams\stream.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <boost\archive\text_iarchive.hpp>
#include "MAPISerializer.hpp"

typedef std::vector<char> buffer_t;

namespace UnitTests
{
/*
	TEST_CLASS(FileSerializationTest)
	{
	public:

		TEST_METHOD(TestPDF)
		{
			char * buffer;
			size_t size = read_file_into_buffer(TESTFILEIN, &buffer);

			std::ofstream ofs;
			ofs.open(TESTFILEOUT, std::ios::binary);
			ofs.write(buffer, size);
			ofs.flush();
			ofs.close();
			
			// TODO: check files are the same
		}
	};
*/

	TEST_CLASS(MAPISerializerTests)
	{
	public:
		
		TEST_METHOD(TestMapiMessage)
		{
			const char SUBJECT[] = "Subject line";
			const char NOTETEXT[] = "This is my message content.\nOn two lines.";
			const char DATERECEIVED[] = "2017/11/30 11:11";

			buffer_t buffer;
			{
				boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
				boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
				boost::archive::text_oarchive ar(s);

				MapiMessage msg;
				msg.ulReserved = 0;
				msg.lpszSubject = new char[100];
				strcpy_s(msg.lpszSubject, 100, SUBJECT);
				msg.lpszNoteText = new char[100];
				strcpy_s(msg.lpszNoteText, 100, NOTETEXT);
				msg.lpszMessageType = NULL;
				msg.lpszDateReceived = new char[strlen(DATERECEIVED) + 1];
				strcpy_s(msg.lpszDateReceived, strlen(DATERECEIVED) + 1, DATERECEIVED);
				msg.lpszConversationID = "";
				msg.flFlags = 0L;
				msg.lpOriginator = NULL; /* Originator descriptor                  */
				msg.nRecipCount = 0;            /* Number of recipients                   */
				msg.lpRecips = NULL;     /* Recipient descriptors                  */
				msg.nFileCount = 0;             /* # of file attachments                  */
				msg.lpFiles = NULL;       /* Attachment descriptors                 */

				ar << msg;
				s.flush();

				Assert::IsTrue(buffer.size() > 0, L"buffer is empty", LINE_INFO());
			}
			{
				boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
				boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
				boost::archive::text_iarchive ia(s);

				MapiMessage * msg = new MapiMessage();
				ia >> *msg;

				Assert::IsTrue(msg->ulReserved == 0, L"ulReserved != 0", LINE_INFO());
				Assert::IsTrue(strcmp(msg->lpszSubject, SUBJECT) == 0, L"lpszSubject != SUBJECT", LINE_INFO());
				Assert::IsTrue(strcmp(msg->lpszNoteText, NOTETEXT) == 0, L"lpszNoteText != NOTETEXT", LINE_INFO());
				Assert::IsNull(msg->lpszMessageType, L"lpszMessageType != NULL", LINE_INFO());
				Assert::IsTrue(strcmp(msg->lpszDateReceived, DATERECEIVED) == 0, L"lpszDateReceived != DATERECEIVED", LINE_INFO());
				Assert::IsTrue(strcmp(msg->lpszConversationID, "") == 0, L"lpszConversationID != \"\"", LINE_INFO());
				Assert::IsTrue(msg->flFlags == 0, L"flFlags != 0", LINE_INFO());
			}
		}

		TEST_METHOD(TestMapiMessageW)
		{
			const WCHAR SUBJECT[] = L"Subject line";
			const WCHAR NOTETEXT[] = L"This is my message content.\nOn three lines.\nSpecial characters ай";
			const WCHAR DATERECEIVED[] = L"2017/11/30 11:11";

			buffer_t buffer;
			{
				boost::iostreams::back_insert_device<buffer_t> inserter(buffer);
				boost::iostreams::stream<boost::iostreams::back_insert_device<buffer_t> > s(inserter);
				boost::archive::text_oarchive ar(s);

				MapiMessageW msg;
				msg.ulReserved = 0;
				msg.lpszSubject = new WCHAR[100];
				wcscpy_s(msg.lpszSubject, 100, SUBJECT);
				msg.lpszNoteText = new WCHAR[100];
				wcscpy_s(msg.lpszNoteText, 100, NOTETEXT);
				msg.lpszMessageType = NULL;
				msg.lpszDateReceived = new WCHAR[wcslen(DATERECEIVED) + 1];
				wcscpy_s(msg.lpszDateReceived, wcslen(DATERECEIVED) + 1, DATERECEIVED);
				msg.lpszConversationID = L"";
				msg.flFlags = 0L;
				msg.lpOriginator = NULL; /* Originator descriptor                  */
				msg.nRecipCount = 0;            /* Number of recipients                   */
				msg.lpRecips = NULL;     /* Recipient descriptors                  */
				msg.nFileCount = 0;             /* # of file attachments                  */
				msg.lpFiles = NULL;       /* Attachment descriptors                 */

				ar << msg;
				s.flush();

				Assert::IsTrue(buffer.size() > 0, L"buffer is empty", LINE_INFO());
			}
			{
				boost::iostreams::basic_array_source<char> device(buffer.data(), buffer.size());
				boost::iostreams::stream<boost::iostreams::basic_array_source<char> > s(device);
				boost::archive::text_iarchive ia(s);

				MapiMessageW msg;
				ia >> msg;

				Assert::IsTrue(msg.ulReserved == 0, L"ulReserved != 0", LINE_INFO());
				Assert::IsTrue(wcscmp(msg.lpszSubject, SUBJECT) == 0, L"lpszSubject != SUBJECT", LINE_INFO());
				Assert::IsTrue(wcscmp(msg.lpszNoteText, NOTETEXT) == 0, L"lpszNoteText != NOTETEXT", LINE_INFO());
				Assert::IsNull(msg.lpszMessageType, L"lpszMessageType != NULL", LINE_INFO());
				Assert::IsTrue(wcscmp(msg.lpszDateReceived, DATERECEIVED) == 0, L"lpszDateReceived != DATERECEIVED", LINE_INFO());
				Assert::IsTrue(wcscmp(msg.lpszConversationID, L"") == 0, L"lpszConversationID != \"\"", LINE_INFO());
				Assert::IsTrue(msg.flFlags == 0, L"flFlags != 0", LINE_INFO());
			}
		}

	};
}