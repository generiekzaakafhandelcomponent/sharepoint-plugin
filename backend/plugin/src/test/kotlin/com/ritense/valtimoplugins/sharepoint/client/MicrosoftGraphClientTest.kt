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

import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody
import com.microsoft.graph.models.DriveItem
import com.microsoft.graph.models.odataerrors.ODataError
import com.microsoft.graph.serviceclient.GraphServiceClient
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Test
import org.mockito.Answers
import org.mockito.kotlin.any
import org.mockito.kotlin.argumentCaptor
import org.mockito.kotlin.mock
import org.mockito.kotlin.verify
import org.mockito.kotlin.whenever
import java.io.ByteArrayInputStream

private class TestODataError(statusCode: Int) : ODataError() {
    init {
        responseStatusCode = statusCode
    }
}

internal class MicrosoftGraphClientTest {
    private val graphServiceClient: GraphServiceClient =
        mock(defaultAnswer = Answers.RETURNS_DEEP_STUBS)
    private val client = MicrosoftGraphClient(graphServiceClient)

    @Test
    fun `createFolder returns the existing folder when Graph reports a conflict`() {
        val conflict = TestODataError(409)
        val existingFolder = DriveItem().apply { id = "existing"; webUrl = "https://sharepoint.example/existing" }

        whenever(graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any()).children().post(any()))
            .thenThrow(conflict)
        whenever(graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any()).get())
            .thenReturn(existingFolder)

        val result = client.createFolder("drive-1", "parent", "folder")

        assertEquals(existingFolder.name, result?.name)
    }

    @Test
    fun `createFolder propagates non-conflict errors`() {
        val forbidden = TestODataError(403)
        whenever(graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any()).children().post(any()))
            .thenThrow(forbidden)

        assertThrows(ODataError::class.java) {
            client.createFolder("drive-1", "parent", "folder")
        }
    }

    @Test
    fun `listDriveItems returns an empty page when the folder is not found`() {
        val notFound = TestODataError(404)
        whenever(graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any()).children().get(any()))
            .thenThrow(notFound)

        val page = client.listDriveItems("drive-1", "missing/folder")

        assertEquals(DriveItemPage(emptyList(), null), page)
    }

    @Test
    fun `listDriveItems propagates non-404 errors`() {
        val serverError = TestODataError(500)
        whenever(graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any()).children().get(any()))
            .thenThrow(serverError)

        assertThrows(ODataError::class.java) {
            client.listDriveItems("drive-1", "some/folder")
        }
    }

    @Test
    fun `listDriveItems rejects a nextLink pointing at a different drive`() {
        assertThrows(IllegalArgumentException::class.java) {
            client.listDriveItems(
                driveId = "drive-1",
                folderPath = "some/folder",
                nextLink = "https://graph.microsoft.com/v1.0/drives/other-drive/items/abc/children?\$skiptoken=x",
            )
        }
    }

    @Test
    fun `listDriveItems rejects a nextLink pointing at a different host`() {
        assertThrows(IllegalArgumentException::class.java) {
            client.listDriveItems(
                driveId = "drive-1",
                folderPath = "some/folder",
                nextLink = "https://attacker.example/v1.0/drives/drive-1/items/abc/children",
            )
        }
    }

    @Test
    fun `listDriveItems rejects a nextLink pointing at a different Graph resource`() {
        assertThrows(IllegalArgumentException::class.java) {
            client.listDriveItems(
                driveId = "drive-1",
                folderPath = "some/folder",
                nextLink = "https://graph.microsoft.com/v1.0/me",
            )
        }
    }

    @Test
    fun `listFolderChildren rejects a nextLink that pages another folder of the same drive`() {
        assertThrows(IllegalArgumentException::class.java) {
            client.listFolderChildren(
                driveId = "drive-1",
                folderItemId = "folder-1",
                nextLink = "https://graph.microsoft.com/v1.0/drives/drive-1/items/other-folder/children?\$skiptoken=x",
            )
        }
    }

    @Test
    fun `uploadFile asks SharePoint to keep both files on a name conflict`() {
        val createUploadSession = graphServiceClient.drives().byDriveId(any()).items().byDriveItemId(any())
            .createUploadSession()
        whenever(createUploadSession.post(any())).thenThrow(TestODataError(500))

        assertThrows(ODataError::class.java) {
            client.uploadFile("drive-1", "folder-1", "brief.docx", ByteArrayInputStream(ByteArray(1)), 1)
        }

        val body = argumentCaptor<CreateUploadSessionPostRequestBody>()
        verify(createUploadSession).post(body.capture())
        assertEquals("rename", body.firstValue.item?.additionalData?.get("@microsoft.graph.conflictBehavior"))
    }
}
