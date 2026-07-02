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

package com.ritense.valtimoplugins.sharepoint.client

import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import io.github.oshai.kotlinlogging.KotlinLogging
import org.springframework.stereotype.Component

private val logger = KotlinLogging.logger {}

@SkipComponentScan
@Component
class MicrosoftGraphClient(
    val graphServiceClient: GraphServiceClient,
) {
    fun listDriveItems(
        driveId: String,
        folderPath: String,
    ): List<DriveItem> {
        return try {
            // "root:/Zaakdossiers/besluit/2026/ZAAK-001:" — colon-terminated path addressing
            val itemId = "root:/${folderPath.trim('/')}:"

            graphServiceClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .children().get()
                ?.value ?: emptyList()
        } catch (e: ODataError) {
            if (e.responseStatusCode == 404) {
                logger.warn { "SharePoint folder not found at path: $folderPath" }
                emptyList()
            } else {
                throw e
            }
        }
    }

    fun createFolder(
        driveId: String,
        parentPath: String,
        folderName: String,
    ): DriveItem? {
        val newFolder = DriveItem().apply {
            name = folderName
            folder = Folder()
            additionalData = mutableMapOf("@microsoft.graph.conflictBehavior" to "rename")
        }

        val parentItemId = if (parentPath.isBlank()) {
            "root"
        } else {
            "root:/${parentPath.trim('/')}:"
        }

        return graphServiceClient
            .drives().byDriveId(driveId)
            .items().byDriveItemId(parentItemId)
            .children().post(newFolder)
    }
}
