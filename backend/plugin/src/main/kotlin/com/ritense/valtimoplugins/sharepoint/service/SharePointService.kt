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

import com.microsoft.graph.models.odataerrors.ODataError
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.service.model.TestConnectionResult
import com.ritense.valtimoplugins.sharepoint.service.model.WorkDocument
import com.ritense.valtimoplugins.sharepoint.service.model.WorkDocumentPage
import io.github.oshai.kotlinlogging.KotlinLogging
import org.springframework.stereotype.Service

private val logger = KotlinLogging.logger {}

@SkipComponentScan
@Service
class SharePointService {
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
        val page = graphClient.listDriveItems(driveId, folderPath, pageSize, nextLink)
        val documents = page.items
            .filter { it.file != null }
            .map { item ->
                WorkDocument(
                    id = item.id ?: "",
                    name = item.name ?: "",
                    webUrl = item.webUrl ?: "",
                    size = item.size,
                    lastModifiedDateTime = item.lastModifiedDateTime?.toString(),
                    createdDateTime = item.createdDateTime?.toString(),
                    thumbnailUrl = item.thumbnails?.firstOrNull()?.medium?.url,
                )
            }
        return WorkDocumentPage(documents, page.nextLink)
    }

    fun createZaakFolder(
        graphClient: MicrosoftGraphClient,
        driveId: String,
        dossierDefinitionName: String,
        year: String,
        zaaknummer: String,
    ): String {
        val parentPath = buildFolderPath( dossierDefinitionName, year)
        return graphClient.createFolder(driveId, parentPath, zaaknummer)?.webUrl ?: ""
    }

    private fun buildFolderPath(vararg parts: String): String =
        parts.filter { it.isNotBlank() }.joinToString("/")
}
