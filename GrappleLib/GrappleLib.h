/*
** Grapple
** Copyright (C) 2005-2010 Will Hui.
**
** Distributed under the terms of the MIT license.
** See LICENSE file for details.
**
** GrappleLib.h
** Exported DLL symbol declarations.
*/

#pragma once

/*
 * The following ifdef block is the standard way of creating macros which make exporting 
 * from a DLL simpler. All files within this DLL are compiled with the GRAPPLELIB_EXPORTS
 * symbol defined on the command line. This symbol should not be defined on any project
 * that uses this DLL. This way any other project whose source files include this file
 * see GRAPPLELIB_API functions as being imported from a DLL, whereas this DLL sees symbols
 * defined with this macro as being exported.
 */
#ifdef GRAPPLELIB_EXPORTS
#define GRAPPLELIB_API __declspec(dllexport)
#else
#define GRAPPLELIB_API __declspec(dllimport)
#endif

// Don't forget to keep GrappleLib.def in sync with this list!
GRAPPLELIB_API bool WINAPI InstallHook(void);
GRAPPLELIB_API void WINAPI RemoveHook(void);
