/*
 * Copyright 2026 Ritense BV, the Netherlands.
 *
 * Licensed under EUPL, Version 1.2 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://joinup.ec.europa.eu/collection/eupl/eupl-text-eupl-12
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.ritense.valtimoplugins.sharepoint.service

import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.odataerrors.ODataError
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.client.DriveItemPage
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.domain.TestConnectionResult
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocument
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentNotFoundException
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentPage
import com.ritense.valtimoplugins.sharepoint.domain.ZaakFolderLocation
import io.github.oshai.kotlinlogging.KotlinLogging
import org.springframework.stereotype.Service
import java.io.InputStream
import java.nio.file.Files
import java.nio.file.StandardCopyOption
import java.util.concurrent.ConcurrentHashMap

private val logger = KotlinLogging.logger {}

@SkipComponentScan
@Service
class SharePointService {

    private val driveIdCache = ConcurrentHashMap<String, String>()

    fun resolveDriveId(
        graphClient: MicrosoftGraphClient,
        hostname: String,
        sharePointSiteName: String,
        baseFolderPath: String,
    ): String {
        return driveIdCache.computeIfAbsent(driveIdCacheKey(hostname, sharePointSiteName, baseFolderPath)) {
            val siteId = graphClient.getSiteId(hostname, sharePointSiteName)
            graphClient.getDriveIdByName(siteId, baseFolderPath)
        }
    }

    fun evictDriveId(hostname: String, sharePointSiteName: String, baseFolderPath: String) {
        driveIdCache.remove(driveIdCacheKey(hostname, sharePointSiteName, baseFolderPath))
    }

    private fun driveIdCacheKey(hostname: String, sharePointSiteName: String, baseFolderPath: String): String =
        "$hostname|$sharePointSiteName|$baseFolderPath"

    fun testConnection(
        tenantId: String,
        clientId: String,
        clientSecret: String,
        hostname: String,
        sharePointSiteName: String,
        baseFolderPath: String,
    ): TestConnectionResult {
        return try {
            val graphClient = MicrosoftGraphClient.build(tenantId, clientId, clientSecret)
            val siteId = graphClient.getSiteId(hostname, sharePointSiteName)
            graphClient.getDriveIdByName(siteId, baseFolderPath)
            TestConnectionResult(success = true, message = "Connection successful.")
        } catch (e: ODataError) {
            logger.warn(e) { "SharePoint connection test failed for $hostname/sites/$sharePointSiteName" }
            TestConnectionResult(
                success = false,
                message = "Could not find the site or document library '$baseFolderPath'. If the hostname " +
                    "and site name look correct, verify that Sites.Read.All (or Sites.Selected) is granted " +
                    "and admin-consented for this app registration — SharePoint returns \"not found\" " +
                    "instead of \"forbidden\" when consent is missing.",
            )
        } catch (e: Exception) {
            logger.warn(e) { "SharePoint connection test failed for $hostname/sites/$sharePointSiteName" }
            TestConnectionResult(success = false, message = "Could not connect to SharePoint: ${e.message}")
        }
    }

    fun listWorkDocuments(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        dossierDefinitionName: String,
        year: String,
        zaaknummer: String,
        pageSize: Int? = null,
        nextLink: String? = null,
    ): WorkDocumentPage {
        val folderPath = buildFolderPath(dossierDefinitionName, year, zaaknummer)
        return toWorkDocumentPage(graphClient.listDriveItems(driveId, folderPath, pageSize, nextLink))
    }

    fun createZaakFolder(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        dossierDefinitionName: String,
        year: String,
        zaaknummer: String,
    ): String {
        val parentPath = buildFolderPath(dossierDefinitionName, year)
        val folderName = validatePathSegment(zaaknummer)
        return graphClient.createFolder(driveId, parentPath, folderName)?.webUrl ?: ""
    }

    fun listZaakWorkDocuments(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        location: ZaakFolderLocation,
        pageSize: Int? = null,
        nextLink: String? = null,
    ): WorkDocumentPage {
        val folder = graphClient.getItemByPath(driveId, zaakFolderPath(location))
            ?: return WorkDocumentPage(emptyList(), null)
        return toWorkDocumentPage(graphClient.listFolderChildren(driveId, requireNotNull(folder.id), pageSize, nextLink))
    }

    fun uploadToZaakFolder(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        location: ZaakFolderLocation,
        fileName: String,
        content: InputStream,
    ): WorkDocument {
        val safeFileName = sanitizeFileName(fileName)
        val folder = graphClient.getItemByPath(driveId, zaakFolderPath(location))
            ?: graphClient.createFolder(driveId, buildFolderPath(location.dossierDefinitionName, location.year), validatePathSegment(location.zaaknummer))
        val folderId = requireNotNull(folder?.id) { "Could not create the SharePoint folder for zaak ${location.zaaknummer}" }

        // The upload session needs the size up front, so the content is buffered to a temporary file
        val tempFile = Files.createTempFile("sharepoint-upload-", ".tmp")
        try {
            content.use { Files.copy(it, tempFile, StandardCopyOption.REPLACE_EXISTING) }
            require(Files.size(tempFile) > 0) { "'$fileName' has no content to upload" }
            val uploaded = Files.newInputStream(tempFile).use {
                graphClient.uploadFile(driveId, folderId, safeFileName, it, Files.size(tempFile))
            }
            return toWorkDocument(uploaded)
        } finally {
            Files.deleteIfExists(tempFile)
        }
    }

    fun getZaakWorkDocument(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        location: ZaakFolderLocation,
        itemId: String,
    ): DriveItem {
        require(ITEM_ID_PATTERN.matches(itemId)) { "Invalid SharePoint item id '$itemId'" }
        val folder = graphClient.getItemByPath(driveId, zaakFolderPath(location))
        val item = folder?.id?.let { graphClient.getItem(driveId, itemId) }
        if (item == null || item.parentReference?.id != folder.id) {
            throw WorkDocumentNotFoundException("No document '$itemId' in the SharePoint folder of zaak ${location.zaaknummer}")
        }
        if (item.file == null) {
            throw IllegalArgumentException("SharePoint item '${item.name}' is not a file")
        }
        return item
    }

    fun downloadContent(graphClient: MicrosoftGraphClient, driveId: String, item: DriveItem): InputStream =
        graphClient.downloadContent(driveId, requireNotNull(item.id))

    fun deleteZaakWorkDocument(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        location: ZaakFolderLocation,
        itemId: String,
    ) {
        val item = getZaakWorkDocument(graphClient, driveId, location, itemId)
        graphClient.deleteItem(driveId, requireNotNull(item.id))
    }

    fun deleteZaakWorkDocumentByName(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        location: ZaakFolderLocation,
        requestedFileName: String,
    ): Boolean {
        // SharePoint strips leading and trailing spaces from names, so a name typed in a form is trimmed too
        val fileName = requestedFileName.trim()
        require(fileName.isNotBlank() && '/' !in fileName && '\\' !in fileName && fileName != "." && fileName != "..") {
            "Invalid SharePoint file name '$fileName'"
        }
        val folderPath = zaakFolderPath(location)
        val item = graphClient.getItemByPath(driveId, "$folderPath/$fileName")
        if (item?.file == null) {
            logger.warn { "No file '$fileName' in SharePoint folder $folderPath; nothing to delete" }
            return false
        }
        graphClient.deleteItem(driveId, requireNotNull(item.id))
        return true
    }

    fun zaakFolderPath(location: ZaakFolderLocation): String =
        buildFolderPath(location.dossierDefinitionName, location.year, location.zaaknummer)

    private fun toWorkDocumentPage(page: DriveItemPage): WorkDocumentPage =
        WorkDocumentPage(page.items.filter { it.file != null }.map(::toWorkDocument), page.nextLink)

    private fun toWorkDocument(item: DriveItem) = WorkDocument(
        id = item.id ?: "",
        name = item.name ?: "",
        webUrl = item.webUrl ?: "",
        size = item.size,
        lastModifiedDateTime = item.lastModifiedDateTime?.toString(),
        createdDateTime = item.createdDateTime?.toString(),
        thumbnailUrl = item.thumbnails?.firstOrNull()?.medium?.url,
    )

    private fun sanitizeFileName(fileName: String): String {
        val sanitized = fileName.replace(FORBIDDEN_FILE_NAME_CHARACTERS, "_").trim().trimEnd('.', ' ')
        require(sanitized.isNotEmpty()) { "Invalid file name '$fileName'" }
        return sanitized
    }

    private fun buildFolderPath(vararg parts: String): String =
        parts.filter { it.isNotBlank() }
            .map(::validatePathSegment)
            .joinToString("/")

    private fun validatePathSegment(segment: String): String {
        require(PATH_SEGMENT_PATTERN.matches(segment)) {
            "Invalid value '$segment': only letters, digits, spaces, hyphens and underscores are allowed."
        }
        return segment
    }

    companion object {
        private val PATH_SEGMENT_PATTERN = Regex("^[\\p{L}\\p{N} _-]+$")
        private val ITEM_ID_PATTERN = Regex("^[A-Za-z0-9!_-]+$")
        private val FORBIDDEN_FILE_NAME_CHARACTERS = Regex("[\"*:<>?/\\\\|]")
    }
}
