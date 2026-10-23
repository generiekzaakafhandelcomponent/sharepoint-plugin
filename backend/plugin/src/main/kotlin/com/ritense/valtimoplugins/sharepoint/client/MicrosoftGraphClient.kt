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

import com.azure.identity.ClientSecretCredentialBuilder
import com.microsoft.graph.core.tasks.LargeFileUploadTask
import com.microsoft.graph.drives.item.items.item.children.ChildrenRequestBuilder
import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody
import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.DriveItemUploadableProperties
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import io.github.oshai.kotlinlogging.KotlinLogging
import java.io.InputStream
import java.net.URI
import java.net.URISyntaxException

private val logger = KotlinLogging.logger {}

private const val GRAPH_HOST = "graph.microsoft.com"

data class DriveItemPage(
    val items: List<DriveItem>,
    val nextLink: String?,
)

class MicrosoftGraphClient(
    val graphServiceClient: GraphServiceClient,
) {
    companion object {
        fun build(tenantId: String, clientId: String, clientSecret: String): MicrosoftGraphClient {
            val credential = ClientSecretCredentialBuilder()
                .tenantId(tenantId)
                .clientId(clientId)
                .clientSecret(clientSecret)
                .build()
            return MicrosoftGraphClient(GraphServiceClient(credential, "https://graph.microsoft.com/.default"))
        }
    }

    fun listDriveItems(
        driveId: String,
        folderPath: String,
        pageSize: Int? = null,
        nextLink: String? = null,
    ): DriveItemPage {
        // "root:/Zaakdossiers/besluit/2026/ZAAK-001:" — colon-terminated path addressing
        return listChildren(driveId, "root:/${folderPath.trim('/')}:", pageSize, nextLink, folderItemId = null)
    }

    fun listFolderChildren(
        driveId: String,
        folderItemId: String,
        pageSize: Int? = null,
        nextLink: String? = null,
    ): DriveItemPage = listChildren(driveId, folderItemId, pageSize, nextLink, folderItemId)

    private fun listChildren(
        driveId: String,
        itemId: String,
        pageSize: Int?,
        nextLink: String?,
        folderItemId: String?,
    ): DriveItemPage {
        return try {
            val response = if (nextLink != null) {
                require(isValidChildrenNextLink(nextLink, driveId, folderItemId)) {
                    "Invalid nextLink for drive '$driveId'"
                }
                ChildrenRequestBuilder(nextLink, graphServiceClient.requestAdapter).get()
            } else {
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
                logger.warn { "SharePoint folder not found: $itemId" }
                DriveItemPage(emptyList(), null)
            } else {
                throw e
            }
        }
    }

    /**
     * Restricts a client-supplied `nextLink` cursor to the children collection of this exact drive,
     * so it cannot be abused to make the server issue authenticated requests to arbitrary Graph
     * endpoints, other tenants/sites, or hosts outside Microsoft Graph.
     */
    private fun isValidChildrenNextLink(nextLink: String, driveId: String, folderItemId: String?): Boolean {
        val uri = try {
            URI(nextLink)
        } catch (e: URISyntaxException) {
            return false
        }
        if (uri.scheme != "https" || uri.host != GRAPH_HOST) {
            return false
        }
        val segments = uri.path.trim('/').split("/")
        // expected: {version}/drives/{driveId}/items/{itemId}/children
        return segments.size == 6 &&
            (segments[0] == "v1.0" || segments[0] == "beta") &&
            segments[1] == "drives" &&
            segments[2] == driveId &&
            segments[3] == "items" &&
            (folderItemId == null || segments[4] == folderItemId) &&
            segments[5] == "children"
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

    fun getItemByPath(driveId: String, path: String): DriveItem? =
        getItem(driveId, "root:/${path.trim('/')}:")

    fun getItem(driveId: String, itemId: String): DriveItem? {
        return try {
            graphServiceClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .get()
        } catch (e: ODataError) {
            if (e.responseStatusCode == 404) null else throw e
        }
    }

    fun downloadContent(driveId: String, itemId: String): InputStream =
        requireNotNull(
            graphServiceClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId(itemId)
                .content().get()
        ) { "No content for SharePoint item $itemId" }

    fun deleteItem(driveId: String, itemId: String) {
        graphServiceClient
            .drives().byDriveId(driveId)
            .items().byDriveItemId(itemId)
            .delete()
    }

    fun uploadFile(
        driveId: String,
        folderItemId: String,
        fileName: String,
        content: InputStream,
        size: Long,
    ): DriveItem {
        val sessionRequest = CreateUploadSessionPostRequestBody().apply {
            item = DriveItemUploadableProperties().apply {
                additionalData = mutableMapOf<String, Any>("@microsoft.graph.conflictBehavior" to "rename")
            }
        }
        val uploadSession = requireNotNull(
            graphServiceClient
                .drives().byDriveId(driveId)
                .items().byDriveItemId("$folderItemId:/$fileName:")
                .createUploadSession().post(sessionRequest)
        ) { "SharePoint did not open an upload session for '$fileName'" }

        val result = LargeFileUploadTask(
            graphServiceClient.requestAdapter,
            uploadSession,
            content,
            size,
            DriveItem::createFromDiscriminatorValue,
        ).upload()
        check(result.isUploadSuccessful) { "Upload of '$fileName' to SharePoint did not complete" }
        return requireNotNull(result.itemResponse) { "SharePoint returned no item for '$fileName'" }
    }
}
