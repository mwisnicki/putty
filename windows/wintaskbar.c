/*
 * wintaskbar.c: contains functions to utilize Windows 7-specific taskbar features.
 *
 * The Windows 7 jumplist is a customizable list defined by the application. It is persistant across application restarts: the OS maintains the list
 *  when the app is not running. The list is shown when the user right-clicks on the taskbar button of a running app or a pinned non-running application.
 *  We use the jumplist to maintain a list of recently started saved sessions, started either by doubleclicking on a saved session, or with
 *  the command line "-load" parameter.
 *
 * Since the jumplist is write-only: it can only be replaced and the current list cannot be read, we must maintain the contents of the list persistantly
 *  in the registry. The file winstore.h contains functions to directly manipulate these registry entries. This file contains higher level functions to manipulate the jumplist.
 *
 */ 

#include "putty.h"
#include "storage.h"
#include "wintaskbar.h"

/* Adds a saved session to the Windows 7 jumplist. */
void add_session_to_jumplist(const char * const sessionname)
{
    if (add_to_jumplist_registry(sessionname) == JUMPLISTREG_OK) {
        update_jumplist_from_registry();
    }
    else {
        /* Make sure we don't leave the jumplist dangling. */
        clear_jumplist();
    }
}


/* Removes a saved session from the Windows jumplist. */
void remove_session_from_jumplist(const char * const sessionname)
{
    if (remove_from_jumplist_registry(sessionname) == JUMPLISTREG_OK) {
        update_jumplist_from_registry();
    }
    else
    {
        /* Make sure we don't leave the jumplist dangling. */
        clear_jumplist();
    }
}

/* Updates jumplist from registry. */
void update_jumplist_from_registry (void)
{
    ICustomDestinationList *pCDL;                        /* ICustomDestinationList is the COM Object used to create application jumplists. */
    char *pjumplist_reg_entries, *piterator;              
    char *param_string, *desc_string;                     
    IPropertyStore *pPS;
    PROPVARIANT pv;
    IObjectCollection *pjumplist_items;                  
    IObjectArray *pjumplist_items_as_array;
    IShellLink *pSL[MAX_JUMPLIST_ITEMS], *pcurrent_link; 
    IObjectArray *pRemoved;                              /* Items removed by user from jump list. These may not be re-added. We don't use this, because IShellLink
                                                               items can't be removed from the jumplist, only pinned. */
    UINT num_items;
    int jumplist_counter;
    HKEY psettings_tmp;
    char b[2048], *p, *q, *r, *putty_path;
    

    /* Retrieve path to Putty executable. */
    GetModuleFileName(NULL, b, sizeof(b) - 1);
    r = b;
    p = strrchr(b, '\\');
    if (p && p >= r) r = p+1;
    q = strrchr(b, ':');
    if (q && q >= r) r = q+1;
    strcpy(r, "putty.exe");
    putty_path = dupstr(b);

    pjumplist_reg_entries = get_jumplist_registry_entries();
    
    if (CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pCDL) == S_OK) {        
        
        if (SUCCEEDED(pCDL->lpVtbl->BeginList(pCDL, &num_items, &IID_IObjectArray, &pRemoved))) {
          
            if (SUCCEEDED(CoCreateInstance(&CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, &IID_IObjectCollection, &pjumplist_items))) {
	            
                /* Walk through the jumplist entries from the registry and add them to the new jumplist. */ 
                piterator = pjumplist_reg_entries;
                jumplist_counter = 0;
                while (*piterator != '\0' && (jumplist_counter < min(MAX_JUMPLIST_ITEMS,(int) num_items)) )
                {
                   /* Check if this is a valid session, otherwise don't add. */
                   psettings_tmp = open_settings_r(piterator);
                   if (psettings_tmp != NULL )
                   {
                       close_settings_r(psettings_tmp);
                         
                       /* Create and add the new item to the list. */
                       if (SUCCEEDED(CoCreateInstance(&CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, &IID_IShellLink, &(pSL[jumplist_counter]) ))) {

                            pcurrent_link = pSL[jumplist_counter];
                            /* Set path, parameters, icon and description. */
                            pcurrent_link->lpVtbl->SetPath(pcurrent_link, putty_path);

                            param_string = snewn(7 + strlen(piterator) + 1 + 1, char);
                            strcpy(param_string, "-load \"");
                            strcat(param_string, piterator);
                            strcat(param_string, "\"");
                            pcurrent_link->lpVtbl->SetArguments(pcurrent_link, param_string);
                            sfree(param_string);

                            desc_string = snewn(20 + strlen(piterator) + 1 + 1, char);
                            strcpy(desc_string, "Connect to session '");
                            strcat(desc_string, piterator);
                            strcat(desc_string, "'");
                            pcurrent_link->lpVtbl->SetDescription(pcurrent_link, desc_string);
                            sfree(desc_string);

                            pcurrent_link->lpVtbl->SetIconLocation(pcurrent_link, putty_path, 0);

                            /* To set the link title, we require the property store of the link. */
                            if (SUCCEEDED(pcurrent_link->lpVtbl->QueryInterface(pcurrent_link, &IID_IPropertyStore, &pPS))) {                    
                                PropVariantInit(&pv);
                                pv.vt = VT_LPSTR;
                                pv.pszVal = dupstr(piterator);
                                pPS->lpVtbl->SetValue(pPS, &PKEY_Title, &pv);
                                sfree(pv.pszVal);
                                pPS->lpVtbl->Commit(pPS);

                                /* Add link to jumplist. */
                                pjumplist_items->lpVtbl->AddObject(pjumplist_items, pcurrent_link);		
                       
	                            pPS->lpVtbl->Release(pPS);
                            }
                            
                            ++jumplist_counter;
                        }            
                    }
                    
                    piterator += strlen(piterator) + 1;
                }
                
                /* Now commit the jumplist. */
                if (SUCCEEDED( pjumplist_items->lpVtbl->QueryInterface(pjumplist_items, &IID_IObjectArray, &pjumplist_items_as_array) )) {
                    pCDL->lpVtbl->AppendCategory(pCDL, L"Recent Sessions", pjumplist_items_as_array);
                    pCDL->lpVtbl->CommitList(pCDL);
                    pjumplist_items_as_array->lpVtbl->Release(pjumplist_items_as_array);
                }

                /* Release COM objects used for the Shell links. */
                --jumplist_counter;
                while (jumplist_counter >= 0)
                {
                    pSL[jumplist_counter]->lpVtbl->Release(pSL[jumplist_counter]);
                    --jumplist_counter;
                }
                
                pjumplist_items->lpVtbl->Release(pjumplist_items);
            }

            pRemoved->lpVtbl->Release(pRemoved);
        }

        pCDL->lpVtbl->Release(pCDL);
    }

    sfree(putty_path);
    sfree(pjumplist_reg_entries);
}

/* Clears the entire jumplist. */
void clear_jumplist(void) {

    ICustomDestinationList *pCDL;
    UINT num_items;         
    IObjectArray *pRemoved; 
    
    if (CoCreateInstance(&CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, &IID_ICustomDestinationList, &pCDL) == S_OK) {
        /* We don't use DeleteList, because that requires an explicitely set AppID. */
        if (SUCCEEDED(pCDL->lpVtbl->BeginList(pCDL, &num_items, &IID_IObjectArray, &pRemoved))){
            pCDL->lpVtbl->CommitList(pCDL);
            pRemoved->lpVtbl->Release(pRemoved);
        }
            
        pCDL->lpVtbl->Release(pCDL);
    }

}
