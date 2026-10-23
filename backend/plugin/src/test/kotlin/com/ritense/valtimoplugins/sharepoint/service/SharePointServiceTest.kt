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
import com.microsoft.graph.models.File
import com.microsoft.graph.models.Folder
import com.microsoft.graph.models.ItemReference
import com.ritense.document.domain.Document
import com.ritense.valtimoplugins.sharepoint.client.DriveItemPage
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentNotFoundException
import com.ritense.valtimoplugins.sharepoint.domain.ZaakFolderLocation
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertFalse
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.Test
import org.mockito.Answers
import org.mockito.kotlin.any
import org.mockito.kotlin.anyOrNull
import org.mockito.kotlin.argumentCaptor
import org.mockito.kotlin.eq
import org.mockito.kotlin.mock
import org.mockito.kotlin.never
import org.mockito.kotlin.verify
import org.mockito.kotlin.whenever
import java.io.ByteArrayInputStream
import java.time.LocalDateTime

internal class SharePointServiceTest {
    private val graphClient: MicrosoftGraphClient = mock()
    private val sharePointService = SharePointService()

    @Test
    fun `listWorkDocuments builds folder path from dossier definition name, year and zaaknummer`() {
        whenever(graphClient.listDriveItems("drive-1", "example/2026/ZAAK-001", 25, null))
            .thenReturn(DriveItemPage(emptyList(), null))

        sharePointService.listWorkDocuments(
            graphClient = graphClient,
            driveId = "drive-1",
            dossierDefinitionName = "example",
            year = "2026",
            zaaknummer = "ZAAK-001",
            pageSize = 25,
        )

        verify(graphClient).listDriveItems("drive-1", "example/2026/ZAAK-001", 25, null)
    }

    @Test
    fun `createZaakFolder builds the parent path without the zaaknummer`() {
        val createdFolder = DriveItem().apply { webUrl = "https://sharepoint.example/folder" }
        whenever(graphClient.createFolder("drive-1", "example/2026", "ZAAK-001")).thenReturn(createdFolder)

        val location = sharePointService.createZaakFolder(
            graphClient = graphClient,
            driveId = "drive-1",
            dossierDefinitionName = "example",
            year = "2026",
            zaaknummer = "ZAAK-001",
        )

        assertEquals("https://sharepoint.example/folder", location)
        verify(graphClient).createFolder("drive-1", "example/2026", "ZAAK-001")
    }

    @Test
    fun `createZaakFolder skips blank path segments`() {
        whenever(graphClient.createFolder("drive-1", "example", "ZAAK-001")).thenReturn(DriveItem())

        sharePointService.createZaakFolder(
            graphClient = graphClient,
            driveId = "drive-1",
            dossierDefinitionName = "example",
            year = "",
            zaaknummer = "ZAAK-001",
        )

        verify(graphClient).createFolder("drive-1", "example", "ZAAK-001")
    }

    @Test
    fun `createZaakFolder rejects path traversal in the zaaknummer`() {
        assertThrows(IllegalArgumentException::class.java) {
            sharePointService.createZaakFolder(
                graphClient = graphClient,
                driveId = "drive-1",
                dossierDefinitionName = "example",
                year = "2026",
                zaaknummer = "ZAAK-001/../../sensitiveFolder",
            )
        }
    }

    @Test
    fun `listWorkDocuments rejects path traversal in the dossier definition name`() {
        assertThrows(IllegalArgumentException::class.java) {
            sharePointService.listWorkDocuments(
                graphClient = graphClient,
                driveId = "drive-1",
                dossierDefinitionName = "../other",
                year = "2026",
                zaaknummer = "ZAAK-001",
            )
        }
    }

    private val location = ZaakFolderLocation("example", "2026", "ZAAK-001")
    private val folder = DriveItem().apply { id = "folder-1"; this.folder = Folder() }

    private fun file(itemId: String, parentId: String = "folder-1", name: String = "brief.docx") = DriveItem().apply {
        id = itemId
        this.name = name
        file = File()
        parentReference = ItemReference().apply { id = parentId }
    }

    @Test
    fun `zaak folder location uses the year the case was created`() {
        val document: Document = mock(defaultAnswer = Answers.RETURNS_DEEP_STUBS)
        whenever(document.definitionId().name()).thenReturn("example")
        whenever(document.createdOn()).thenReturn(LocalDateTime.of(2025, 12, 31, 23, 0))

        assertEquals(ZaakFolderLocation("example", "2025", "ZAAK-001"), ZaakFolderLocation.of(document, "ZAAK-001"))
    }

    @Test
    fun `listZaakWorkDocuments lists the files of the zaak folder by its id`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.listFolderChildren("drive-1", "folder-1", 25, null))
            .thenReturn(DriveItemPage(listOf(file("item-1"), DriveItem().apply { id = "sub"; this.folder = Folder() }), "next"))

        val page = sharePointService.listZaakWorkDocuments(graphClient, "drive-1", location, 25)

        assertEquals(listOf("item-1"), page.documents.map { it.id })
        assertEquals("next", page.nextLink)
    }

    @Test
    fun `listZaakWorkDocuments returns an empty page when the zaak folder does not exist`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(null)

        val page = sharePointService.listZaakWorkDocuments(graphClient, "drive-1", location)

        assertTrue(page.documents.isEmpty())
        verify(graphClient, never()).listFolderChildren(any(), any(), anyOrNull(), anyOrNull())
    }

    @Test
    fun `uploadToZaakFolder uploads the whole content into the zaak folder under a name SharePoint accepts`() {
        val content = "document content".toByteArray()
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.uploadFile(eq("drive-1"), eq("folder-1"), any(), any(), any()))
            .thenReturn(file("item-2", name = "brief_ concept_.docx"))

        val uploaded = sharePointService.uploadToZaakFolder(
            graphClient, "drive-1", location, " brief: concept?.docx. ", ByteArrayInputStream(content)
        )

        val size = argumentCaptor<Long>()
        verify(graphClient).uploadFile(eq("drive-1"), eq("folder-1"), eq("brief_ concept_.docx"), any(), size.capture())
        assertEquals(content.size.toLong(), size.firstValue)
        assertEquals("item-2", uploaded.id)
    }

    @Test
    fun `uploadToZaakFolder refuses a document without content`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)

        assertThrows(IllegalArgumentException::class.java) {
            sharePointService.uploadToZaakFolder(graphClient, "drive-1", location, "leeg.pdf", ByteArrayInputStream(ByteArray(0)))
        }
        verify(graphClient, never()).uploadFile(any(), any(), any(), any(), any())
    }

    @Test
    fun `uploadToZaakFolder creates the zaak folder when it does not exist yet`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(null)
        whenever(graphClient.createFolder("drive-1", "example/2026", "ZAAK-001")).thenReturn(folder)
        whenever(graphClient.uploadFile(any(), any(), any(), any(), any())).thenReturn(file("item-2"))

        sharePointService.uploadToZaakFolder(graphClient, "drive-1", location, "brief.docx", ByteArrayInputStream(ByteArray(1)))

        verify(graphClient).uploadFile(eq("drive-1"), eq("folder-1"), eq("brief.docx"), any(), any())
    }

    @Test
    fun `getZaakWorkDocument refuses an item that is not in the zaak folder`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.getItem("drive-1", "item-9")).thenReturn(file("item-9", parentId = "other-folder"))

        assertThrows(WorkDocumentNotFoundException::class.java) {
            sharePointService.getZaakWorkDocument(graphClient, "drive-1", location, "item-9")
        }
    }

    @Test
    fun `getZaakWorkDocument refuses an item id that addresses a path`() {
        assertThrows(IllegalArgumentException::class.java) {
            sharePointService.getZaakWorkDocument(graphClient, "drive-1", location, "root:/other/brief.docx:")
        }
        verify(graphClient, never()).getItem(any(), any())
    }

    @Test
    fun `getZaakWorkDocument refuses a folder`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.getItem("drive-1", "sub")).thenReturn(
            DriveItem().apply { id = "sub"; this.folder = Folder(); parentReference = ItemReference().apply { id = "folder-1" } }
        )

        assertThrows(IllegalArgumentException::class.java) {
            sharePointService.getZaakWorkDocument(graphClient, "drive-1", location, "sub")
        }
    }

    @Test
    fun `deleteZaakWorkDocument deletes a file in the zaak folder`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.getItem("drive-1", "item-1")).thenReturn(file("item-1"))

        sharePointService.deleteZaakWorkDocument(graphClient, "drive-1", location, "item-1")

        verify(graphClient).deleteItem("drive-1", "item-1")
    }

    @Test
    fun `deleteZaakWorkDocument deletes nothing outside the zaak folder`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001")).thenReturn(folder)
        whenever(graphClient.getItem("drive-1", "item-9")).thenReturn(file("item-9", parentId = "other-folder"))

        assertThrows(WorkDocumentNotFoundException::class.java) {
            sharePointService.deleteZaakWorkDocument(graphClient, "drive-1", location, "item-9")
        }
        verify(graphClient, never()).deleteItem(any(), any())
    }

    @Test
    fun `deleteZaakWorkDocumentByName deletes the named file in the zaak folder`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001/brief.docx")).thenReturn(file("item-1"))

        assertTrue(sharePointService.deleteZaakWorkDocumentByName(graphClient, "drive-1", location, "brief.docx"))

        verify(graphClient).deleteItem("drive-1", "item-1")
    }

    @Test
    fun `deleteZaakWorkDocumentByName ignores spaces around a name typed in a form`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001/brief.docx")).thenReturn(file("item-1"))

        assertTrue(sharePointService.deleteZaakWorkDocumentByName(graphClient, "drive-1", location, " brief.docx "))

        verify(graphClient).deleteItem("drive-1", "item-1")
    }

    @Test
    fun `deleteZaakWorkDocumentByName completes without deleting when the file is already gone`() {
        whenever(graphClient.getItemByPath("drive-1", "example/2026/ZAAK-001/brief.docx")).thenReturn(null)

        assertFalse(sharePointService.deleteZaakWorkDocumentByName(graphClient, "drive-1", location, "brief.docx"))

        verify(graphClient, never()).deleteItem(any(), any())
    }

    @Test
    fun `deleteZaakWorkDocumentByName rejects a name that leaves the zaak folder`() {
        listOf("../brief.docx", "sub/brief.docx", "..", " ").forEach { name ->
            assertThrows(IllegalArgumentException::class.java) {
                sharePointService.deleteZaakWorkDocumentByName(graphClient, "drive-1", location, name)
            }
        }
        verify(graphClient, never()).deleteItem(any(), any())
    }
}
