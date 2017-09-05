#include "CWMPEventDispatch.h"

HRESULT CWMPEventDispatch::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS FAR* pDispParams, VARIANT FAR*, EXCEPINFO FAR*, unsigned int FAR*)
{
	if (!pDispParams)
		return E_POINTER;

	if (pDispParams->cNamedArgs != 0)
		return DISP_E_NONAMEDARGS;

	return DISP_E_MEMBERNOTFOUND;
}

void CWMPEventDispatch::OpenStateChange(long)
{
}

void CWMPEventDispatch::PlayStateChange(long NewState)
{
	if (NewState == 0x3)
		PostMessage(_hMainWindowHandle, WM_APP, 0, 0);
}

void CWMPEventDispatch::AudioLanguageChange(long)
{
}

void CWMPEventDispatch::StatusChange()
{
}

void CWMPEventDispatch::ScriptCommand(BSTR, BSTR)
{
}

void CWMPEventDispatch::NewStream()
{
}

void CWMPEventDispatch::Disconnect(long)
{
}

void CWMPEventDispatch::Buffering(VARIANT_BOOL)
{
}

void CWMPEventDispatch::Error()
{
}

void CWMPEventDispatch::Warning(long, long, BSTR)
{
}

void CWMPEventDispatch::EndOfStream(long)
{
}

void CWMPEventDispatch::PositionChange(double, double)
{
}

void CWMPEventDispatch::MarkerHit(long)
{
}

void CWMPEventDispatch::DurationUnitChange(long)
{
}

void CWMPEventDispatch::CdromMediaChange(long)
{
}

void CWMPEventDispatch::PlaylistChange(IDispatch *, WMPPlaylistChangeEventType)
{
}

void CWMPEventDispatch::CurrentPlaylistChange(WMPPlaylistChangeEventType)
{
}

void CWMPEventDispatch::CurrentPlaylistItemAvailable(BSTR)
{
}

void CWMPEventDispatch::MediaChange(IDispatch *)
{
}

void CWMPEventDispatch::CurrentMediaItemAvailable(BSTR)
{
}

void CWMPEventDispatch::CurrentItemChange(IDispatch *)
{
}

void CWMPEventDispatch::MediaCollectionChange()
{
}

void CWMPEventDispatch::MediaCollectionAttributeStringAdded(BSTR, BSTR)
{
}

void CWMPEventDispatch::MediaCollectionAttributeStringRemoved(BSTR, BSTR)
{
}

void CWMPEventDispatch::MediaCollectionAttributeStringChanged(BSTR, BSTR, BSTR)
{
}

void CWMPEventDispatch::PlaylistCollectionChange()
{
}

void CWMPEventDispatch::PlaylistCollectionPlaylistAdded(BSTR)
{
}

void CWMPEventDispatch::PlaylistCollectionPlaylistRemoved(BSTR)
{
}

void CWMPEventDispatch::PlaylistCollectionPlaylistSetAsDeleted(BSTR, VARIANT_BOOL)
{
}

void CWMPEventDispatch::ModeChange(BSTR, VARIANT_BOOL)
{
}

void CWMPEventDispatch::MediaError(IDispatch *)
{
}

void CWMPEventDispatch::OpenPlaylistSwitch(IDispatch *)
{
}

void CWMPEventDispatch::DomainChange(BSTR)
{
}

void CWMPEventDispatch::SwitchedToPlayerApplication()
{
}

void CWMPEventDispatch::SwitchedToControl()
{
}

void CWMPEventDispatch::PlayerDockedStateChange()
{
}

void CWMPEventDispatch::PlayerReconnect()
{
}

void CWMPEventDispatch::Click(short, short, long, long)
{
}

void CWMPEventDispatch::DoubleClick(short, short, long, long)
{
}

void CWMPEventDispatch::KeyDown(short, short)
{
}

void CWMPEventDispatch::KeyPress(short)
{
}

void CWMPEventDispatch::KeyUp(short, short)
{
}

void CWMPEventDispatch::MouseDown(short, short, long, long)
{
}

void CWMPEventDispatch::MouseMove(short, short, long, long)
{
}

void CWMPEventDispatch::MouseUp(short, short, long, long)
{
}
