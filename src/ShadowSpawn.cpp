/* 
Copyright (c) 2011 Craig Andera (shadowspawn@wangdera.com)

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/


#include "stdafx.h"
#include "CComException.h"
#include "CShadowSpawnException.h"
#include "OutputWriter.h"
#include "CWriter.h"
#include "CWriterComponent.h"
#include "Exports.h"



// Forward declarations
void CalculateSourcePath(LPCTSTR wszSnapshotDevice, LPCTSTR wszBackupSource, LPCTSTR wszMountPoint, CString& output);
bool ShouldAddComponent(CWriterComponent& component);



void CalculateSourcePath(LPCTSTR wszSnapshotDevice, LPCTSTR wszBackupSource, LPCTSTR wszMountPoint, CString& output)
{
	CString backupSource(wszBackupSource); 
	CString mountPoint(wszMountPoint); 

	CString subdirectory = backupSource.Mid(mountPoint.GetLength()); 

	Utilities::CombinePath(wszSnapshotDevice, subdirectory, output); 
}

void Cleanup(bool bAbnormalAbort, bool bSnapshotCreated, const CString& mountedDevice, CComPtr<IVssBackupComponents> pBackupComponents, GUID snapshotSetId,OutputWriter& logger)
{
	if (pBackupComponents == NULL)
	{
		return; 
	}

	if (bAbnormalAbort)
	{
		logger.WriteLine(TEXT("Aborting backup."), VERBOSITY_THRESHOLD_NORMAL);
		pBackupComponents->AbortBackup(); 
	}
	if (!mountedDevice.IsEmpty())
	{
		if (bAbnormalAbort)
		{
			CString message;
			message.AppendFormat(TEXT("Dismounting device: %s"), mountedDevice);
			logger.WriteLine(message, VERBOSITY_THRESHOLD_NORMAL);
		}
		BOOL bWorked = DefineDosDevice(DDD_REMOVE_DEFINITION, mountedDevice, NULL); 
		if (!bWorked)
		{
			DWORD error = ::GetLastError(); 
			CString errorMessage; 
			Utilities::FormatErrorMessage(error, errorMessage); 
			CString message; 
			message.AppendFormat(TEXT("There was an error calling DefineDosDevice during Cleanup. Error: %s"), errorMessage); 
			logger.WriteLine(message);
		}
	}
	if (bSnapshotCreated)
	{
		if (bAbnormalAbort)
		{
			logger.WriteLine(TEXT("Deleting snapshot."), VERBOSITY_THRESHOLD_NORMAL);
		}
		LONG cDeletedSnapshots; 
		GUID nonDeletedSnapshotId; 
		pBackupComponents->DeleteSnapshots(snapshotSetId, VSS_OBJECT_SNAPSHOT_SET, TRUE, 
			&cDeletedSnapshots, &nonDeletedSnapshotId); 
	}
}

bool ShouldAddComponent(CWriterComponent& component)
{
	// Component should not be added if
	// 1) It is not selectable for backup and 
	// 2) It has a selectable ancestor
	// Otherwise, add it. 

	if (component.get_SelectableForBackup())
	{
		return true; 
	}

	return !component.get_HasSelectableAncestor();

}


GUID GetSystemProviderID(OutputWriter& logger)
{
	CComPtr<IVssBackupComponents> backupComponents; 

	logger.WriteLine(TEXT("Calling CreateVssBackupComponents in GetSystemProviderId")); 
	CHECK_HRESULT(::CreateVssBackupComponents(&backupComponents)); 

	logger.WriteLine(TEXT("Calling InitializeForBackup in GetSystemProviderId")); 
	CHECK_HRESULT(backupComponents->InitializeForBackup()); 

	// The following code for selecting the system proviider is necessary 
	// per http://forum.storagecraft.com/Community/forums/p/177/542.aspx#542
	// which is a totally awesome post
	logger.WriteLine(TEXT("Looking for the system VSS provider"));

	logger.WriteLine(TEXT("Calling backupComponents->Query(enum providers)"));
	CComPtr<IVssEnumObject> pEnum; 
	CHECK_HRESULT(backupComponents->Query(GUID_NULL, VSS_OBJECT_NONE, VSS_OBJECT_PROVIDER, &pEnum));

	GUID systemProviderId = GUID_NULL;
	VSS_OBJECT_PROP prop;
	ULONG nFetched;
	do 
	{
		logger.WriteLine(TEXT("Calling IVssEnumObject::Next"));
		HRESULT hr = pEnum->Next(1, &prop, &nFetched);

		CString message;
		message.AppendFormat(TEXT("Examining provider %s to see if it's the system provider..."), prop.Obj.Prov.m_pwszProviderName);
		logger.WriteLine(message);

		if (hr == S_OK)
		{
			if (prop.Obj.Prov.m_eProviderType == VSS_PROV_SYSTEM)
			{
				systemProviderId = prop.Obj.Prov.m_ProviderId; 
				logger.WriteLine(TEXT("...and it is."));
				break;
			}
		}
		else if (hr == S_FALSE)
		{
			logger.WriteLine(TEXT("...but it's not."));
			break;
		}
		else
		{
			throw new CComException(hr, __FILE__, __LINE__);
		}
	} while (true);

	if (systemProviderId.Data1 == GUID_NULL.Data1 && 
		systemProviderId.Data2 == GUID_NULL.Data2 &&
		systemProviderId.Data3 == GUID_NULL.Data3 &&
		systemProviderId.Data4 == GUID_NULL.Data4)
	{
		throw new CShadowSpawnException(TEXT("Unable to locate the system snapshot provider."));
	}

	return systemProviderId;
}

HRESULT _ShadowSpawn(LPCTSTR source,LPCTSTR device,bool debug,int verbosityLevel,bool simulate,ShadowSpawnCallback* callback,LogCallback* logCallback)
{
	OutputWriter logger;
	logger.SetLogger(logCallback);

	bool bSnapshotCreated = false;
	CString mountedDevice;
	CComPtr<IVssBackupComponents> pBackupComponents; 
	GUID snapshotSetId = GUID_NULL; 

	int fileCount = 0; 
	LONGLONG byteCount = 0; 
	int directoryCount = 0; 
	int skipCount = 0; 
	SYSTEMTIME startTime;
	try
	{

		if (debug)
		{
			::DebugBreak(); 
		}

		logger.SetVerbosityLevel((VERBOSITY_LEVEL) verbosityLevel); 

		if (!Utilities::DirectoryExists(source))
		{
			CString message;
			message.AppendFormat(TEXT("Source path is not an existing directory: %s"), source);
			throw new CShadowSpawnException(HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND),message); 
		}



		::GetSystemTime(&startTime); 
		CString startTimeString; 
		Utilities::FormatDateTime(&startTime, TEXT(" "), false, startTimeString); 

		CString message; 
		message.AppendFormat(TEXT("Shadowing %s at %s"), 
			source, 
			device); 
		logger.WriteLine(message, VERBOSITY_THRESHOLD_NORMAL); 

		GUID systemProviderId = GetSystemProviderID(logger);

		logger.WriteLine(TEXT("Calling CreateVssBackupComponents")); 
		CHECK_HRESULT(::CreateVssBackupComponents(&pBackupComponents)); 

		logger.WriteLine(TEXT("Calling InitializeForBackup")); 
		CHECK_HRESULT(pBackupComponents->InitializeForBackup()); 

		CComPtr<IVssAsync> pWriterMetadataStatus; 

		logger.WriteLine(TEXT("Calling GatherWriterMetadata")); 
		CHECK_HRESULT(pBackupComponents->GatherWriterMetadata(&pWriterMetadataStatus)); 

		logger.WriteLine(TEXT("Waiting for call to GatherWriterMetadata to finish...")); 
		CHECK_HRESULT(pWriterMetadataStatus->Wait()); 

		HRESULT hrGatherStatus; 
		logger.WriteLine(TEXT("Calling QueryStatus for GatherWriterMetadata")); 
		CHECK_HRESULT(pWriterMetadataStatus->QueryStatus(&hrGatherStatus, NULL)); 

		if (hrGatherStatus == VSS_S_ASYNC_CANCELLED)
		{
			throw new CShadowSpawnException(L"GatherWriterMetadata was cancelled."); 
		}

		logger.WriteLine(TEXT("Call to GatherWriterMetadata finished.")); 


		logger.WriteLine(TEXT("Calling GetWriterMetadataCount")); 

		vector<CWriter> writers;

		UINT cWriters; 
		CHECK_HRESULT(pBackupComponents->GetWriterMetadataCount(&cWriters)); 

		for (UINT iWriter = 0; iWriter < cWriters; ++iWriter)
		{
			CWriter writer; 
			CComPtr<IVssExamineWriterMetadata> pExamineWriterMetadata; 
			GUID id; 
			logger.WriteLine(TEXT("Calling GetWriterMetadata")); 
			CHECK_HRESULT(pBackupComponents->GetWriterMetadata(iWriter, &id, &pExamineWriterMetadata)); 
			GUID idInstance; 
			GUID idWriter; 
			BSTR bstrWriterName;
			VSS_USAGE_TYPE usage; 
			VSS_SOURCE_TYPE source; 
			CHECK_HRESULT(pExamineWriterMetadata->GetIdentity(&idInstance, &idWriter, &bstrWriterName, &usage, &source)); 

			writer.set_InstanceId(idInstance); 
			writer.set_Name(bstrWriterName); 
			writer.set_WriterId(idWriter); 

			CComBSTR writerName(bstrWriterName); 
			CString message; 
			message.AppendFormat(TEXT("Writer %d named %s"), iWriter, (LPCTSTR) writerName); 
			logger.WriteLine(message); 

			UINT cIncludeFiles;
			UINT cExcludeFiles; 
			UINT cComponents; 
			CHECK_HRESULT(pExamineWriterMetadata->GetFileCounts(&cIncludeFiles, &cExcludeFiles, &cComponents)); 

			message.Empty(); 
			message.AppendFormat(TEXT("Writer has %d components"), cComponents); 
			logger.WriteLine(message); 

			for (UINT iComponent = 0; iComponent < cComponents; ++iComponent)
			{
				CWriterComponent component; 

				CComPtr<IVssWMComponent> pComponent; 
				CHECK_HRESULT(pExamineWriterMetadata->GetComponent(iComponent, &pComponent)); 

				PVSSCOMPONENTINFO pComponentInfo; 
				CHECK_HRESULT(pComponent->GetComponentInfo(&pComponentInfo)); 

				CString message; 
				message.AppendFormat(TEXT("Component %d is named %s, has a path of %s, and is %sselectable for backup. %d files, %d databases, %d log files."), 
					iComponent,
					pComponentInfo->bstrComponentName, 
					pComponentInfo->bstrLogicalPath, 
					pComponentInfo->bSelectable ? TEXT("") : TEXT("not "), 
					pComponentInfo->cFileCount, 
					pComponentInfo->cDatabases,
					pComponentInfo->cLogFiles); 
				logger.WriteLine(message); 

				component.set_LogicalPath(pComponentInfo->bstrLogicalPath); 
				component.set_SelectableForBackup(pComponentInfo->bSelectable); 
				component.set_Writer(iWriter); 
				component.set_Name(pComponentInfo->bstrComponentName);
				component.set_Type(pComponentInfo->type);

				for (UINT iFile = 0; iFile < pComponentInfo->cFileCount; ++iFile)
				{
					CComPtr<IVssWMFiledesc> pFileDesc; 
					CHECK_HRESULT(pComponent->GetFile(iFile, &pFileDesc)); 

					CComBSTR bstrPath; 
					CHECK_HRESULT(pFileDesc->GetPath(&bstrPath)); 

					CComBSTR bstrFileSpec; 
					CHECK_HRESULT(pFileDesc->GetFilespec(&bstrFileSpec)); 

					CString message; 
					message.AppendFormat(TEXT("File %d has path %s\\%s"), iFile, bstrPath, bstrFileSpec); 
					logger.WriteLine(message); 
				}

				for (UINT iDatabase = 0; iDatabase < pComponentInfo->cDatabases; ++iDatabase)
				{
					CComPtr<IVssWMFiledesc> pFileDesc; 
					CHECK_HRESULT(pComponent->GetDatabaseFile(iDatabase, &pFileDesc)); 

					CComBSTR bstrPath; 
					CHECK_HRESULT(pFileDesc->GetPath(&bstrPath)); 

					CComBSTR bstrFileSpec; 
					CHECK_HRESULT(pFileDesc->GetFilespec(&bstrFileSpec)); 

					CString message; 
					message.AppendFormat(TEXT("Database file %d has path %s\\%s"), iDatabase, bstrPath, bstrFileSpec); 
					logger.WriteLine(message); 
				}

				for (UINT iDatabaseLogFile = 0; iDatabaseLogFile < pComponentInfo->cLogFiles; ++iDatabaseLogFile)
				{
					CComPtr<IVssWMFiledesc> pFileDesc; 
					CHECK_HRESULT(pComponent->GetDatabaseLogFile(iDatabaseLogFile, &pFileDesc)); 

					CComBSTR bstrPath; 
					CHECK_HRESULT(pFileDesc->GetPath(&bstrPath)); 

					CComBSTR bstrFileSpec; 
					CHECK_HRESULT(pFileDesc->GetFilespec(&bstrFileSpec)); 

					CString message; 
					message.AppendFormat(TEXT("Database log file %d has path %s\\%s"), iDatabaseLogFile, bstrPath, bstrFileSpec); 
					logger.WriteLine(message); 
				}

				CHECK_HRESULT(pComponent->FreeComponentInfo(pComponentInfo)); 

				writer.get_Components().push_back(component); 

			}

			writer.ComputeComponentTree(); 

			for (unsigned int iComponent = 0; iComponent < writer.get_Components().size(); ++iComponent)
			{
				CWriterComponent& component = writer.get_Components()[iComponent]; 
				CString message; 
				message.AppendFormat(TEXT("Component %d has name %s, path %s, is %sselectable for backup, and has parent %s"), 
					iComponent, 
					component.get_Name(), 
					component.get_LogicalPath(), 
					component.get_SelectableForBackup() ? TEXT("") : TEXT("not "), 
					component.get_Parent() == NULL ? TEXT("(no parent)") : component.get_Parent()->get_Name()); 
				logger.WriteLine(message); 
			}

			writers.push_back(writer); 
		}

		logger.WriteLine(TEXT("Calling StartSnapshotSet")); 
		CHECK_HRESULT(pBackupComponents->StartSnapshotSet(&snapshotSetId));

		logger.WriteLine(TEXT("Calling GetVolumePathName")); 
		WCHAR wszVolumePathName[MAX_PATH]; 
		BOOL bWorked = ::GetVolumePathName(source, wszVolumePathName, MAX_PATH); 

		if (!bWorked)
		{
			DWORD error = ::GetLastError(); 
			CString errorMessage; 
			Utilities::FormatErrorMessage(error, errorMessage); 
			CString message; 
			message.AppendFormat(TEXT("There was an error retrieving the volume name from the path. Path: %s Error: %s"), 
				source, errorMessage); 
			throw new CShadowSpawnException(message.GetString()); 
		}


		logger.WriteLine(TEXT("Calling AddToSnapshotSet")); 
		GUID snapshotId; 
		CHECK_HRESULT(pBackupComponents->AddToSnapshotSet(wszVolumePathName, systemProviderId, &snapshotId)); 

		for (unsigned int iWriter = 0; iWriter < writers.size(); ++iWriter)
		{
			CWriter writer = writers[iWriter];

			CString message; 
			message.AppendFormat(TEXT("Adding components to snapshot set for writer %s"), writer.get_Name()); 
			logger.WriteLine(message); 
			for (unsigned int iComponent = 0; iComponent < writer.get_Components().size(); ++iComponent)
			{
				CWriterComponent component = writer.get_Components()[iComponent];

				if (ShouldAddComponent(component))
				{
					CString message; 
					message.AppendFormat(TEXT("Adding component %s (%s) from writer %s"), 
						component.get_Name(), 
						component.get_LogicalPath(), 
						writer.get_Name()); 
					logger.WriteLine(message); 
					CHECK_HRESULT(pBackupComponents->AddComponent(
						writer.get_InstanceId(), 
						writer.get_WriterId(),
						component.get_Type(), 
						component.get_LogicalPath(), 
						component.get_Name()
						));
				}
				else
				{
					CString message; 
					message.AppendFormat(TEXT("Not adding component %s from writer %s."), 
						component.get_Name(), writer.get_Name()); 
					logger.WriteLine(message); 
				}
			}
		}

		logger.WriteLine(TEXT("Calling SetBackupState")); 
		CHECK_HRESULT(pBackupComponents->SetBackupState(TRUE, FALSE, VSS_BACKUP_TYPE::VSS_BT_FULL, FALSE)); 

		logger.WriteLine(TEXT("Calling PrepareForBackup")); 
		CComPtr<IVssAsync> pPrepareForBackupResults; 
		CHECK_HRESULT(pBackupComponents->PrepareForBackup(&pPrepareForBackupResults)); 

		logger.WriteLine(TEXT("Waiting for call to PrepareForBackup to finish...")); 
		CHECK_HRESULT(pPrepareForBackupResults->Wait()); 

		HRESULT hrPrepareForBackupResults; 
		CHECK_HRESULT(pPrepareForBackupResults->QueryStatus(&hrPrepareForBackupResults, NULL)); 

		if (hrPrepareForBackupResults != VSS_S_ASYNC_FINISHED)
		{
			throw new CShadowSpawnException(TEXT("Prepare for backup failed.")); 
		}

		logger.WriteLine(TEXT("Call to PrepareForBackup finished.")); 

		SYSTEMTIME snapshotTime; 
		::GetSystemTime(&snapshotTime); 

		if (!simulate)
		{
			logger.WriteLine(TEXT("Calling DoSnapshotSet")); 
			CComPtr<IVssAsync> pDoSnapshotSetResults;
			CHECK_HRESULT(pBackupComponents->DoSnapshotSet(&pDoSnapshotSetResults)); 

			logger.WriteLine(TEXT("Waiting for call to DoSnapshotSet to finish...")); 

			CHECK_HRESULT(pDoSnapshotSetResults->Wait());

			bSnapshotCreated = true; 

			HRESULT hrDoSnapshotSetResults; 
			CHECK_HRESULT(pDoSnapshotSetResults->QueryStatus(&hrDoSnapshotSetResults, NULL)); 

			if (hrDoSnapshotSetResults != VSS_S_ASYNC_FINISHED)
			{
				throw new CShadowSpawnException(L"DoSnapshotSet failed."); 
			}

			logger.WriteLine(TEXT("Call to DoSnapshotSet finished.")); 

			logger.WriteLine(TEXT("Calling GetSnapshotProperties")); 
			VSS_SNAPSHOT_PROP snapshotProperties; 
			CHECK_HRESULT(pBackupComponents->GetSnapshotProperties(snapshotId, &snapshotProperties));

			logger.WriteLine(TEXT("Calling CalculateSourcePath")); 
			// TODO: We'll eventually have to deal with mount points
			CString wszSource;
			CalculateSourcePath(
				snapshotProperties.m_pwszSnapshotDeviceObject, 
				source,
				wszVolumePathName, 
				wszSource
				);

			logger.WriteLine(TEXT("Calling DefineDosDevice to mount device.")); 
			if (0 == wszSource.Find(TEXT("\\\\?\\GLOBALROOT")))
			{
				wszSource = wszSource.Mid(_tcslen(TEXT("\\\\?\\GLOBALROOT")));
			}
			bWorked = DefineDosDevice(DDD_RAW_TARGET_PATH, device, wszSource); 
			if (!bWorked)
			{
				DWORD error = ::GetLastError(); 
				CString errorMessage; 
				Utilities::FormatErrorMessage(error, errorMessage); 
				CString message; 
				message.AppendFormat(TEXT("There was an error calling DefineDosDevice when mounting a device. Error: %s"), errorMessage); 
				throw new CShadowSpawnException(message.GetString()); 
			}
			mountedDevice = device;

			callback();

			logger.WriteLine(TEXT("Calling DefineDosDevice to remove device.")); 
			bWorked = DefineDosDevice(DDD_REMOVE_DEFINITION, device, NULL); 
			if (!bWorked)
			{
				DWORD error = ::GetLastError(); 
				CString errorMessage; 
				Utilities::FormatErrorMessage(error, errorMessage); 
				CString message; 
				message.AppendFormat(TEXT("There was an error calling DefineDosDevice. Error: %s"), errorMessage); 
				throw new CShadowSpawnException(message.GetString()); 
			}
			mountedDevice.Empty();

			logger.WriteLine(TEXT("Calling BackupComplete")); 
			CComPtr<IVssAsync> pBackupCompleteResults; 
			CHECK_HRESULT(pBackupComponents->BackupComplete(&pBackupCompleteResults)); 

			logger.WriteLine(TEXT("Waiting for call to BackupComplete to finish...")); 
			CHECK_HRESULT(pBackupCompleteResults->Wait());

			HRESULT hrBackupCompleteResults; 
			CHECK_HRESULT(pBackupCompleteResults->QueryStatus(&hrBackupCompleteResults, NULL)); 

			if (hrBackupCompleteResults != VSS_S_ASYNC_FINISHED)
			{
				throw new CShadowSpawnException(TEXT("Completion of backup failed.")); 
			}

			logger.WriteLine(TEXT("Call to BackupComplete finished.")); 

		}
	}
	catch (CComException* e)
	{
		Cleanup(true, bSnapshotCreated, mountedDevice, pBackupComponents, snapshotSetId,logger);
		CString message; 
		CString file; 
		e->get_File(file); 
		message.Format(TEXT("There was a COM failure 0x%x - %s (%d)"), 
			e->get_Hresult(), file, e->get_Line()); 
		logger.WriteLine(message, VERBOSITY_THRESHOLD_UNLESS_SILENT); 
		return e->get_Hresult(); 
	}
	catch (CShadowSpawnException* e)
	{
		Cleanup(true, bSnapshotCreated, mountedDevice, pBackupComponents, snapshotSetId,logger);
		logger.WriteLine(e->get_Message(), VERBOSITY_THRESHOLD_UNLESS_SILENT); 
		return e->get_HResult(); 
	}

	Cleanup(false, bSnapshotCreated, mountedDevice, pBackupComponents, snapshotSetId,logger);
	logger.WriteLine(TEXT("Shadowing successfully completed."), VERBOSITY_THRESHOLD_NORMAL); 
	return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __cdecl ShadowSpawn(LPCTSTR source,LPCTSTR device,int verbosityLevel,ShadowSpawnCallback* callback,LogCallback* logCallback)
{
	return _ShadowSpawn(source,device,false,verbosityLevel,false,callback,logCallback);
}