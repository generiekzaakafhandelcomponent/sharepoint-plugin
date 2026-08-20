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

import com.microsoft.graph.drives.item.items.item.children.ChildrenRequestBuilder
import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import io.github.oshai.kotlinlogging.KotlinLogging

private val logger = KotlinLogging.logger {}

data class DriveItemPage(
    val items: List<DriveItem>,
    val nextLink: String?,
)

class MicrosoftGraphClient(
    val graphServiceClient: GraphServiceClient,
) {
    fun listDriveItems(
        driveId: String,
        folderPath: String,
        pageSize: Int? = null,
        nextLink: String? = null,
    ): DriveItemPage {
        return try {
            val response = if (nextLink != null) {
                ChildrenRequestBuilder(nextLink, graphServiceClient.requestAdapter).get()
            } else {
                // "root:/Zaakdossiers/besluit/2026/ZAAK-001:" — colon-terminated path addressing
                val itemId = "root:/${folderPath.trim('/')}:"

                graphServiceClient
                    .drives().byDriveId(driveId)
                    .items().byDriveItemId(itemId)
                    .children().get { requestConfiguration ->
                        if (pageSize != null) {
                            requestConfiguration.queryParameters.top = pageSize
                        }
                        requestConfiguration.queryParameters.expand = arrayOf("thumbnails")
                    }
            }
            DriveItemPage(response?.value ?: emptyList(), response?.odataNextLink)
        } catch (e: ODataError) {
            if (e.responseStatusCode == 404) {
                logger.warn { "SharePoint folder not found at path: $folderPath" }
                DriveItemPage(emptyList(), null)
            } else {
                throw e
            }
        }
    }

    fun getSiteId(hostname: String, sitePath: String): String =
        requireNotNull(
            graphServiceClient
                .sites()
                .bySiteId("$hostname:/sites/$sitePath")
                .get()
                ?.id
        ) { "Site not found: $hostname/sites/$sitePath" }

    fun getDriveIdByName(siteId: String, driveName: String): String {
        val drive = graphServiceClient
            .sites()
            .bySiteId(siteId)
            .drives()
            .get()
            ?.value
            ?.firstOrNull { it.name.equals(driveName, ignoreCase = true) }
        return requireNotNull(drive?.id) { "Drive '$driveName' not found in site $siteId" }
    }

    fun createFolder(
        driveId: String,
        parentPath: String,
        folderName: String,
    ): DriveItem? {
        val newFolder = DriveItem().apply {
            name = folderName
            folder = Folder()
            additionalData = mutableMapOf("@microsoft.graph.conflictBehavior" to "fail")
        }

        val parentItemId = if (parentPath.isBlank()) "root" else "root:/${parentPath.trim('/')}:"

        return try {
            graphServiceClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(parentItemId)
                .children().post(newFolder)
        } catch (e: ODataError) {
            if (e.responseStatusCode == 409) {
                // Folder already exists — fetch and return the existing one
                val folderPath = if (parentPath.isBlank()) folderName else "${parentPath.trim('/')}/$folderName"
                graphServiceClient
                    .drives().byDriveId(driveId)
                    .items().byDriveItemId("root:/${folderPath}:")
                    .get()
            } else {
                throw e
            }
        }
    }
}
