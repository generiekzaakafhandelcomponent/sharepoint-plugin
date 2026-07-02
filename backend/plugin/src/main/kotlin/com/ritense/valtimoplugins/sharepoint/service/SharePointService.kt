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

import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.service.model.WorkDocument
import org.springframework.stereotype.Service

@SkipComponentScan
@Service
class SharePointService(
    private val graphClient: MicrosoftGraphClient,
) {
    fun listWorkDocuments(
        siteId: String,
        driveId: String,
        baseFolderPath: String,
        zaaktype: String,
        year: String,
        zaaknummer: String,
    ): List<WorkDocument> {
        val folderPath = buildFolderPath(baseFolderPath, zaaktype, year, zaaknummer)
        return graphClient.listDriveItems(siteId, driveId, folderPath)
            .filter { it.file != null }
            .map { item ->
                WorkDocument(
                    id = item.id ?: "",
                    name = item.name ?: "",
                    webUrl = item.webUrl ?: "",
                    size = item.size,
                    lastModifiedDateTime = item.lastModifiedDateTime?.toString(),
                    createdDateTime = item.createdDateTime?.toString(),
                )
            }
    }

    fun createZaakFolder(
        siteId: String,
        driveId: String,
        baseFolderPath: String,
        zaaktype: String,
        year: String,
        zaaknummer: String,
    ) {
        val parentPath = buildFolderPath(baseFolderPath, zaaktype, year)
        graphClient.createFolder(siteId, driveId, parentPath, zaaknummer)
    }

    private fun buildFolderPath(vararg parts: String): String =
        parts.filter { it.isNotBlank() }.joinToString("/")
}
