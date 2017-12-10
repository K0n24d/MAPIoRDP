#pragma once

#include <MAPI.h>
#include <boost\archive\basic_archive.hpp>
#include <boost\serialization\version.hpp>
#include <boost\serialization\split_free.hpp>
#include <boost\serialization\binary_object.hpp>
#include <fstream>
#include <sstream>
#include <string>
#include <boost/filesystem.hpp>

DWORD OpenDynamicChannel(HANDLE *phFile);

template <typename MapiMessageT>
void cleanup_files(const MapiMessageT & msg)
{
	for (ULONG i = 0; i < msg.nFileCount; i++)
		boost::filesystem::remove(boost::filesystem::path(msg.lpFiles[i].lpszPathName));
}

template <typename CharT>
std::size_t read_file_into_buffer(const CharT * lpszfileName, char ** buffer)
{
	std::ifstream ifs;
	ifs.open(lpszfileName, std::ios::binary);
	ifs.seekg(0, std::ios::end);
	std::size_t pos = (std::size_t) ifs.tellg();

	*buffer = (char *)malloc(pos * sizeof(char));
	ifs.seekg(std::ios::beg);
	ifs.read(*buffer, pos);

	return pos;
}

namespace boost {
	namespace serialization {

		template<class Archive>
		void save_c_str(Archive & ar, const char * src)
		{
			size_t size = 0;
			if(src != NULL)
				size = strlen(src) + 1;
			ar & size;
			if(size > 0)
				ar & make_binary_object(src, size);
		}
		template<class Archive>
		void load_c_str(Archive & ar, LPSTR & dest)
		{
			size_t size;
			ar & size;
			if (size > 0)
			{
				dest = (LPSTR)malloc(size * sizeof(CHAR));
				ar & make_binary_object(dest, size);
				(dest)[size-1] = '\0';
			}
			else
				dest = NULL;
		}

		template<class Archive>
		void save(Archive & ar, const MapiRecipDesc & recip, const unsigned int version)
		{
			ar & recip.ulReserved;                                /* Reserved for future use                  */
			ar & recip.ulRecipClass;                              /* Recipient class                          */
																  /* MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG    */
			save_c_str<Archive>(ar, recip.lpszName);              /* Recipient name                           */
			save_c_str<Archive>(ar, recip.lpszAddress);           /* Recipient address (optional)             */
			ar & recip.ulEIDSize;                                 /* Count in bytes of size of pEntryID       */
			ar & make_binary_object(recip.lpEntryID, recip.ulEIDSize); /* System-specific recipient reference      */
		}
		template<class Archive>
		void load(Archive & ar, MapiRecipDesc & recip, const unsigned int version)
		{
			ar & recip.ulReserved;                                     /* Reserved for future use                  */
			ar & recip.ulRecipClass;                                   /* Recipient class                          */
																	   /* MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG    */
			load_c_str<Archive>(ar, recip.lpszName);                   /* Recipient name                           */
			load_c_str<Archive>(ar, recip.lpszAddress);                /* Recipient address (optional)             */
			ar & recip.ulEIDSize;                                      /* Count in bytes of size of pEntryID       */
			ar & make_binary_object(recip.lpEntryID, recip.ulEIDSize); /* System-specific recipient reference      */
		}

		template<class Archive>
		void save(Archive & ar, const MapiFileTagExt & tag, const unsigned int version)
		{
			ar & tag.ulReserved;                                     /* Reserved, must be zero.                  */
			ar & tag.cbTag;                                          /* Size (in bytes) of                       */
			ar & make_binary_object(tag.lpTag, tag.cbTag);           /* X.400 OID for this attachment type       */
			ar & tag.cbEncoding;                                     /* Size (in bytes) of                       */
			ar & make_binary_object(tag.lpEncoding, tag.cbEncoding); /* X.400 OID for this attachment's encoding */
		}
		template<class Archive>
		void load(Archive & ar, MapiFileTagExt & tag, const unsigned int version)
		{
			ar & tag.ulReserved;                                     /* Reserved, must be zero.                  */
			ar & tag.cbTag;                                          /* Size (in bytes) of                       */
			tag.lpTag = (LPBYTE)malloc(tag.cbTag * sizeof(BYTE));
			ar & make_binary_object(tag.lpTag, tag.cbTag);           /* X.400 OID for this attachment type       */
			ar & tag.cbEncoding;                                     /* Size (in bytes) of                       */
			tag.lpEncoding = (LPBYTE)malloc(tag.cbEncoding * sizeof(BYTE));
			ar & make_binary_object(tag.lpEncoding, tag.cbEncoding); /* X.400 OID for this attachment's encoding */
		}

		template<class Archive>
		void save(Archive & ar, const MapiFileDesc & file, const unsigned int version)
		{
			ar & file.ulReserved;                         /* Reserved for future use (must be 0)     */
			ar & file.flFlags;                            /* Flags                                   */
			ar & file.nPosition;                          /* character in text to be replaced by attachment */

			// file.lpszPathName                          /* Full path name of attachment file       */
			char * fileContent;
			std::size_t size = read_file_into_buffer(file.lpszPathName, &fileContent);
			ar & (ULONG)size;
			ar & make_binary_object(fileContent, size);
			delete fileContent;

			// file.lpszFileName                          /* Original file name (optional)           */
			if(file.lpszFileName!=NULL && strlen(file.lpszFileName)>0)
				save_c_str<Archive>(ar, file.lpszFileName);   
			else
			{
				save_c_str<Archive>(ar, boost::filesystem::path(file.lpszFileName).filename().string().c_str());
			}

			bool hasFileType(file.lpFileType != NULL);
			ar & hasFileType;
			if (hasFileType)
				ar & (lpMapiFileTagExt)(file.lpFileType); /* Attachment file type (can be lpMapiFileTagExt) */
		}
		template<class Archive>
		void load(Archive & ar, MapiFileDesc & file, const unsigned int version)
		{
			ar & file.ulReserved;                         /* Reserved for future use (must be 0)     */
			ar & file.flFlags;                            /* Flags                                   */
			ar & file.nPosition;                          /* character in text to be replaced by attachment */
			

			// file.lpszPathName                          /* Full path name of attachment file       */
			std::size_t size;
			ar & size;
			char * fileContent = (char *)malloc(size * sizeof(char));
			ar & make_binary_object(fileContent, size);
			boost::filesystem::path temp = boost::filesystem::temp_directory_path() / boost::filesystem::unique_path();
			std::ofstream ofs;
			ofs.open(temp.string(), std::ios::binary);
			ofs.write(fileContent, size);
			ofs.flush();
			ofs.close();
			free(fileContent);

			std::string tempstr(temp.string());
			file.lpszPathName = (LPSTR)malloc((tempstr.size() + 1) * sizeof(CHAR));
			strcpy_s(file.lpszPathName, tempstr.size() + 1, tempstr.c_str());

			load_c_str<Archive>(ar, file.lpszFileName);   /* Original file name (optional)           */

			bool hasFileType;
			ar & hasFileType;
			if (hasFileType)
			{
				lpMapiFileTagExt mapiFileTagExt = (lpMapiFileTagExt)malloc(sizeof(MapiFileTagExt));
				ar & mapiFileTagExt; 
				file.lpFileType = mapiFileTagExt;         /* Attachment file type (can be lpMapiFileTagExt) */
			}
			else
				file.lpFileType = NULL;
		}

		template<class Archive>
		void save(Archive & ar, const MapiMessage & msg, const unsigned int version)
		{
			ar & msg.ulReserved;                             /* Reserved for future use (M.B. 0)       */
			save_c_str<Archive>(ar, msg.lpszSubject);        /* Message Subject                        */
			save_c_str<Archive>(ar, msg.lpszNoteText);       /* Message Text                           */
			save_c_str<Archive>(ar, msg.lpszMessageType);    /* Message Class                          */
			save_c_str<Archive>(ar, msg.lpszDateReceived);   /* in YYYY/MM/DD HH:MM format             */
			save_c_str<Archive>(ar, msg.lpszConversationID); /* conversation thread ID                 */
			ar & msg.flFlags;                                /* unread,return receipt                  */
			ar & msg.lpOriginator;                           /* Originator descriptor                  */

			ar & msg.nRecipCount;                            /* Number of recipients                   */
			for (ULONG i = 0; i < msg.nRecipCount; i++)
				ar & msg.lpRecips[i];                        /* Recipient descriptors                  */

			ar & msg.nFileCount;                             /* # of file attachments                  */
			for (ULONG i = 0; i < msg.nFileCount; i++)
				ar & msg.lpFiles[i];                         /* Attachment descriptors                 */
		}
		template<class Archive>
		void load(Archive & ar, MapiMessage & msg, const unsigned int version)
		{
			ar & msg.ulReserved;                             /* Reserved for future use (M.B. 0)       */
			load_c_str<Archive>(ar, msg.lpszSubject);        /* Message Subject                        */
			load_c_str<Archive>(ar, msg.lpszNoteText);       /* Message Text                           */
			load_c_str<Archive>(ar, msg.lpszMessageType);    /* Message Class                          */
			load_c_str<Archive>(ar, msg.lpszDateReceived);   /* in YYYY/MM/DD HH:MM format             */
			load_c_str<Archive>(ar, msg.lpszConversationID); /* conversation thread ID                 */
			ar & msg.flFlags;                                /* unread,return receipt                  */
			ar & msg.lpOriginator;                           /* Originator descriptor                  */

			ar & msg.nRecipCount;                            /* Number of recipients                   */
			msg.lpRecips = (lpMapiRecipDesc)calloc(msg.nRecipCount, sizeof(MapiRecipDesc));
			for (ULONG i = 0; i < msg.nRecipCount; i++)
				ar & msg.lpRecips[i];                        /* Recipient descriptors                  */

			ar & msg.nFileCount;                             /* # of file attachments                  */
			msg.lpFiles = (lpMapiFileDesc)calloc(msg.nFileCount, sizeof(MapiFileDesc));
			for (ULONG i = 0; i < msg.nFileCount; i++)
				ar & msg.lpFiles[i];                         /* Attachment descriptors                 */
		}

		template<class Archive>
		void save_c_str(Archive & ar, const wchar_t * src)
		{
			size_t size = 0;
			if (src != NULL)
			{
				size = wcslen(src) + 1;
				size *= 2;
			}
			ar & size;
			if (size > 0)
				ar & make_binary_object(src, size);
		}
		template<class Archive>
		void load_c_str(Archive & ar, PWSTR & dest)
		{
			size_t size;
			ar & size;
			if (size > 0)
			{
				dest = (PWSTR)malloc(size * sizeof(char));
				ar & make_binary_object(dest, size);
			}
			else
				dest = NULL;
		}

		template<class Archive>
		void save(Archive & ar, const MapiRecipDescW & recip, const unsigned int version)
		{
			ar & recip.ulReserved;                                /* Reserved for future use                  */
			ar & recip.ulRecipClass;                              /* Recipient class                          */
																  /* MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG    */
			save_c_str<Archive>(ar, recip.lpszName);              /* Recipient name                           */
			save_c_str<Archive>(ar, recip.lpszAddress);           /* Recipient address (optional)             */
			ar & recip.ulEIDSize;                                 /* Count in bytes of size of pEntryID       */
			ar & make_binary_object(recip.lpEntryID, recip.ulEIDSize); /* System-specific recipient reference      */
		}
		template<class Archive>
		void load(Archive & ar, MapiRecipDescW & recip, const unsigned int version)
		{
			ar & recip.ulReserved;                                     /* Reserved for future use                  */
			ar & recip.ulRecipClass;                                   /* Recipient class                          */
																	   /* MAPI_TO, MAPI_CC, MAPI_BCC, MAPI_ORIG    */
			load_c_str<Archive>(ar, recip.lpszName);                   /* Recipient name                           */
			load_c_str<Archive>(ar, recip.lpszAddress);                /* Recipient address (optional)             */
			ar & recip.ulEIDSize;                                      /* Count in bytes of size of pEntryID       */
			ar & make_binary_object(recip.lpEntryID, recip.ulEIDSize); /* System-specific recipient reference      */
		}

		template<class Archive>
		void save(Archive & ar, const MapiFileDescW & file, const unsigned int version)
		{
			ar & file.ulReserved;                         /* Reserved for future use (must be 0)     */
			ar & file.flFlags;                            /* Flags                                   */
			ar & file.nPosition;                          /* character in text to be replaced by attachment */

			// file.lpszPathName                          /* Full path name of attachment file       */
			char * fileContent;
			std::size_t size = read_file_into_buffer(file.lpszPathName, &fileContent);
			ar & size;
			ar & make_binary_object(fileContent, size);
			delete fileContent;

			// file.lpszFileName                          /* Original file name (optional)           */
			if (file.lpszFileName != NULL && wcslen(file.lpszFileName)>0)
				save_c_str<Archive>(ar, file.lpszFileName);
			else
				save_c_str<Archive>(ar, boost::filesystem::path(file.lpszFileName).filename().wstring().c_str());

			bool hasFileType(file.lpFileType != NULL);
			ar & hasFileType;
			if (hasFileType)
				ar & (lpMapiFileTagExt)(file.lpFileType); /* Attachment file type (can be lpMapiFileTagExt) */
		}
		template<class Archive>
		void load(Archive & ar, MapiFileDescW & file, const unsigned int version)
		{
			ar & file.ulReserved;                         /* Reserved for future use (must be 0)     */
			ar & file.flFlags;                            /* Flags                                   */
			ar & file.nPosition;                          /* character in text to be replaced by attachment */

			// file.lpszPathName                          /* Full path name of attachment file       */
			std::size_t size;
			ar & size;
			char * fileContent = (char *)malloc(size * sizeof(char));
			ar & make_binary_object(fileContent, size);
			boost::filesystem::path temp = boost::filesystem::temp_directory_path() / boost::filesystem::unique_path();
			std::ofstream ofs;
			ofs.open(temp.wstring(), std::ios::binary);
			ofs.write(fileContent, size);
			ofs.flush();
			ofs.close();
			free(fileContent);

			std::wstring tempstr(temp.wstring());
			file.lpszPathName = (PWSTR)malloc((tempstr.length() + 1) * sizeof(WCHAR));
			wcscpy_s(file.lpszPathName, tempstr.length() + 1, tempstr.c_str());

			load_c_str<Archive>(ar, file.lpszFileName);   /* Original file name (optional)           */

			bool hasFileType;
			ar & hasFileType;
			if (hasFileType)
			{
				lpMapiFileTagExt mapiFileTagExt = (lpMapiFileTagExt)malloc(sizeof(MapiFileTagExt));
				ar & mapiFileTagExt;
				file.lpFileType = mapiFileTagExt;         /* Attachment file type (can be lpMapiFileTagExt) */
			}
			else
				file.lpFileType = NULL;
		}

		template<class Archive>
		void save(Archive & ar, const MapiMessageW & msg, const unsigned int version)
		{
			ar & msg.ulReserved;                             /* Reserved for future use (M.B. 0)       */
			save_c_str<Archive>(ar, msg.lpszSubject);        /* Message Subject                        */
			save_c_str<Archive>(ar, msg.lpszNoteText);       /* Message Text                           */
			save_c_str<Archive>(ar, msg.lpszMessageType);    /* Message Class                          */
			save_c_str<Archive>(ar, msg.lpszDateReceived);   /* in YYYY/MM/DD HH:MM format             */
			save_c_str<Archive>(ar, msg.lpszConversationID); /* conversation thread ID                 */
			ar & msg.flFlags;                                /* unread,return receipt                  */
			ar & msg.lpOriginator;                           /* Originator descriptor                  */

			ar & msg.nRecipCount;                            /* Number of recipients                   */
			for (ULONG i = 0; i < msg.nRecipCount; i++)
				ar & msg.lpRecips[i];                        /* Recipient descriptors                  */

			ar & msg.nFileCount;                             /* # of file attachments                  */
			for (ULONG i = 0; i < msg.nFileCount; i++)
				ar & msg.lpFiles[i];                         /* Attachment descriptors                 */
		}
		template<class Archive>
		void load(Archive & ar, MapiMessageW & msg, const unsigned int version)
		{
			ar & msg.ulReserved;                             /* Reserved for future use (M.B. 0)       */
			load_c_str<Archive>(ar, msg.lpszSubject);        /* Message Subject                        */
			load_c_str<Archive>(ar, msg.lpszNoteText);       /* Message Text                           */
			load_c_str<Archive>(ar, msg.lpszMessageType);    /* Message Class                          */
			load_c_str<Archive>(ar, msg.lpszDateReceived);   /* in YYYY/MM/DD HH:MM format             */
			load_c_str<Archive>(ar, msg.lpszConversationID); /* conversation thread ID                 */
			ar & msg.flFlags;                                /* unread,return receipt                  */
			ar & msg.lpOriginator;                           /* Originator descriptor                  */

			ar & msg.nRecipCount;                            /* Number of recipients                   */
			msg.lpRecips = (lpMapiRecipDescW)calloc(msg.nRecipCount, sizeof(MapiRecipDescW));
			for (ULONG i = 0; i < msg.nRecipCount; i++)
				ar & msg.lpRecips[i];                        /* Recipient descriptors                  */

			ar & msg.nFileCount;                             /* # of file attachments                  */
			msg.lpFiles = (lpMapiFileDescW)calloc(msg.nFileCount, sizeof(MapiFileDescW));
			for (ULONG i = 0; i < msg.nFileCount; i++)
				ar & msg.lpFiles[i];                         /* Attachment descriptors                 */
		}
	} // namespace serialization
} // namespace boost

BOOST_SERIALIZATION_SPLIT_FREE(MapiRecipDesc)
BOOST_SERIALIZATION_SPLIT_FREE(MapiFileTagExt)
BOOST_SERIALIZATION_SPLIT_FREE(MapiFileDesc)
BOOST_SERIALIZATION_SPLIT_FREE(MapiMessage)

BOOST_SERIALIZATION_SPLIT_FREE(MapiRecipDescW)
BOOST_SERIALIZATION_SPLIT_FREE(MapiFileDescW)
BOOST_SERIALIZATION_SPLIT_FREE(MapiMessageW)
