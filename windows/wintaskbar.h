/*
 * Wintaskbar.h: Definitions of preferences and COM structures needed for Windows 7 jumplist manipulation.
 */

#ifndef PUTTY_WINTASKBAR_H
#define PUTTY_WINTASKBAR_H

#define MAX_JUMPLIST_ITEMS 30    /* PuTTY will never show more items in the jumplist than this, regardless of user preferences. */

/*
 * COM structures and functions.
 */

#define IID_IShellLink IID_IShellLinkA

extern const CLSID CLSID_DestinationList;
extern const CLSID CLSID_ShellLink;
extern const CLSID CLSID_EnumerableObjectCollection;
extern const IID IID_IObjectCollection;
extern const IID IID_IShellLink;
extern const IID IID_ICustomDestinationList;
extern const IID IID_IObjectArray;
extern const IID IID_IPropertyStore;
extern const PROPERTYKEY PKEY_Title;


typedef struct ICustomDestinationListVtbl {  
    HRESULT ( __stdcall *QueryInterface ) ( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [in] */  const GUID * const riid,
        /* [out] */ void **ppvObject);
    
    ULONG ( __stdcall *AddRef )(
        /* [in] ICustomDestinationList*/ void *This);
    
    ULONG ( __stdcall *Release )( 
        /* [in] ICustomDestinationList*/ void *This);
    
    HRESULT ( __stdcall *SetAppID )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [string][in] */ LPCWSTR pszAppID);
    
    HRESULT ( __stdcall *BeginList )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [out] */ UINT *pcMinSlots,
        /* [in] */  const GUID * const riid,
        /* [out] */ void **ppv);
    
    HRESULT ( __stdcall *AppendCategory )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [string][in] */ LPCWSTR pszCategory,
        /* [in] IObjectArray*/ void *poa);
    
    HRESULT ( __stdcall *AppendKnownCategory )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [in] KNOWNDESTCATEGORY*/ int category);
    
    HRESULT ( __stdcall *AddUserTasks )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [in] IObjectArray*/ void *poa);
    
    HRESULT ( __stdcall *CommitList )( 
        /* [in] ICustomDestinationList*/ void *This);
    
    HRESULT ( __stdcall *GetRemovedDestinations )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [in] */ const IID * const riid,
        /* [out] */ void **ppv);
    
    HRESULT ( __stdcall *DeleteList )( 
        /* [in] ICustomDestinationList*/ void *This,
        /* [string][unique][in] */ LPCWSTR pszAppID);
    
    HRESULT ( __stdcall *AbortList )( 
        /* [in] ICustomDestinationList*/ void *This);
    
} ICustomDestinationListVtbl;

typedef struct ICustomDestinationList
{
    ICustomDestinationListVtbl *lpVtbl;
} ICustomDestinationList;

typedef struct IObjectArrayVtbl
{
    HRESULT ( __stdcall *QueryInterface )( 
        /* [in] IObjectArray*/ void *This,
        /* [in] */ const GUID * const riid,
        /* [out] */ void **ppvObject);
    
    ULONG ( __stdcall *AddRef )( 
        /* [in] IObjectArray*/ void *This);
    
    ULONG ( __stdcall *Release )( 
        /* [in] IObjectArray*/ void *This);
    
    HRESULT ( __stdcall *GetCount )( 
        /* [in] IObjectArray*/ void *This,
        /* [out] */ UINT *pcObjects);
    
    HRESULT ( __stdcall *GetAt )( 
        /* [in] IObjectArray*/ void *This,
        /* [in] */ UINT uiIndex,
        /* [in] */ const GUID * const riid,
        /* [out] */ void **ppv);

} IObjectArrayVtbl;

typedef struct IObjectArray
{
    IObjectArrayVtbl *lpVtbl;
} IObjectArray;

typedef struct IShellLinkVtbl
{
    HRESULT ( __stdcall *QueryInterface )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ const GUID * const riid,
        /* [out] */ void **ppvObject);
    
    ULONG ( __stdcall *AddRef )( 
        /* [in] IShellLink*/ void *This);
    
    ULONG ( __stdcall *Release )( 
        /* [in] IShellLink*/ void *This);
    
    HRESULT ( __stdcall *GetPath )( 
        /* [in] IShellLink*/ void *This,
        /* [string][out] */ LPSTR pszFile,
        /* [in] */ int cch,
        /* [unique][out][in] */ WIN32_FIND_DATAA *pfd,
        /* [in] */ DWORD fFlags);
    
    HRESULT ( __stdcall *GetIDList )( 
        /* [in] IShellLink*/ void *This,
        /* [out] LPITEMIDLIST*/ void **ppidl);
    
    HRESULT ( __stdcall *SetIDList )( 
        /* [in] IShellLink*/ void *This,
        /* [in] LPITEMIDLIST*/ void *pidl);
    
    HRESULT ( __stdcall *GetDescription )( 
        /* [in] IShellLink*/ void *This,
        /* [string][out] */ LPSTR pszName,
        /* [in] */ int cch);
    
    HRESULT ( __stdcall *SetDescription )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszName);
    
    HRESULT ( __stdcall *GetWorkingDirectory )( 
        /* [in] IShellLink*/ void *This,
        /* [string][out] */ LPSTR pszDir,
        /* [in] */ int cch);
    
    HRESULT ( __stdcall *SetWorkingDirectory )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszDir);
    
    HRESULT ( __stdcall *GetArguments )( 
        /* [in] IShellLink*/ void *This,
        /* [string][out] */ LPSTR pszArgs,
        /* [in] */ int cch);
    
    HRESULT ( __stdcall *SetArguments )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszArgs);
    
    HRESULT ( __stdcall *GetHotkey )( 
        /* [in] IShellLink*/ void *This,
        /* [out] */ WORD *pwHotkey);
    
    HRESULT ( __stdcall *SetHotkey )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ WORD wHotkey);
    
    HRESULT ( __stdcall *GetShowCmd )( 
        /* [in] IShellLink*/ void *This,
        /* [out] */ int *piShowCmd);
    
    HRESULT ( __stdcall *SetShowCmd )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ int iShowCmd);
    
    HRESULT ( __stdcall *GetIconLocation )( 
        /* [in] IShellLink*/ void *This,
        /* [string][out] */ LPSTR pszIconPath,
        /* [in] */ int cch,
        /* [out] */ int *piIcon);
    
    HRESULT ( __stdcall *SetIconLocation )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszIconPath,
        /* [in] */ int iIcon);
    
    HRESULT ( __stdcall *SetRelativePath )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszPathRel,
        /* [in] */ DWORD dwReserved);
    
    HRESULT ( __stdcall *Resolve )( 
        /* [in] IShellLink*/ void *This,
        /* [unique][in] */ HWND hwnd,
        /* [in] */ DWORD fFlags);
    
    HRESULT ( __stdcall *SetPath )( 
        /* [in] IShellLink*/ void *This,
        /* [string][in] */ LPCSTR pszFile);
    
} IShellLinkVtbl;

typedef struct IShellLink
{
    IShellLinkVtbl *lpVtbl;
} IShellLink;

typedef struct IObjectCollectionVtbl
{
    HRESULT ( __stdcall *QueryInterface )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ const GUID * const riid,
        /* [out] */ void **ppvObject);
    
    ULONG ( __stdcall *AddRef )( 
        /* [in] IShellLink*/ void *This);
    
    ULONG ( __stdcall *Release )( 
        /* [in] IShellLink*/ void *This);
    
    HRESULT ( __stdcall *GetCount )( 
        /* [in] IShellLink*/ void *This,
        /* [out] */ UINT *pcObjects);
    
    HRESULT ( __stdcall *GetAt )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ UINT uiIndex,
        /* [in] */ const GUID * const riid,
        /* [iid_is][out] */ void **ppv);
    
    HRESULT ( __stdcall *AddObject )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ void *punk);
    
    HRESULT ( __stdcall *AddFromArray )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ IObjectArray *poaSource);
    
    HRESULT ( __stdcall *RemoveObjectAt )( 
        /* [in] IShellLink*/ void *This,
        /* [in] */ UINT uiIndex);
    
    HRESULT ( __stdcall *Clear )( 
        /* [in] IShellLink*/ void *This);
   
} IObjectCollectionVtbl;

typedef struct IObjectCollection
{
    IObjectCollectionVtbl *lpVtbl;
} IObjectCollection;

typedef struct IPropertyStoreVtbl
{
    HRESULT ( __stdcall *QueryInterface )( 
        /* [in] IPropertyStore*/ void *This,
        /* [in] */ const GUID * const riid,
        /* [iid_is][out] */ void **ppvObject);
    
    ULONG ( __stdcall *AddRef )( 
        /* [in] IPropertyStore*/ void *This);
    
    ULONG ( __stdcall *Release )( 
        /* [in] IPropertyStore*/ void *This);
    
    HRESULT ( __stdcall *GetCount )( 
        /* [in] IPropertyStore*/ void *This,
        /* [out] */ DWORD *cProps);
    
    HRESULT ( __stdcall *GetAt )( 
        /* [in] IPropertyStore*/ void *This,
        /* [in] */ DWORD iProp,
        /* [out] */ PROPERTYKEY *pkey);
    
    HRESULT ( __stdcall *GetValue )( 
        /* [in] IPropertyStore*/ void *This,
        /* [in] */ const PROPERTYKEY * const key,
        /* [out] */ PROPVARIANT *pv);
    
    HRESULT ( __stdcall *SetValue )( 
        /* [in] IPropertyStore*/ void *This,
        /* [in] */ const PROPERTYKEY * const key,
        /* [in] */ REFPROPVARIANT propvar);
    
    HRESULT ( __stdcall *Commit )( 
        /* [in] IPropertyStore*/ void *This);
} IPropertyStoreVtbl;

typedef struct IPropertyStore
{
    IPropertyStoreVtbl *lpVtbl;
} IPropertyStore;

#endif
