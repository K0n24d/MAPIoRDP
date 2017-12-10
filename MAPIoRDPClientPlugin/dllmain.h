// dllmain.h : Déclaration de classe de module.

class CMAPIoRDPClientPluginModule : public ATL::CAtlDllModuleT< CMAPIoRDPClientPluginModule >
{
public :
	DECLARE_LIBID(LIBID_MAPIoRDPClientPluginLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_MAPIORDPCLIENTPLUGIN, "{FCDDA6DD-D90E-441C-B445-5808F6C61882}")
};

extern class CMAPIoRDPClientPluginModule _AtlModule;
