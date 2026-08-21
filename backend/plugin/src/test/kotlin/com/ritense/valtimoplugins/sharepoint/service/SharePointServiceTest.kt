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
import com.ritense.valtimoplugins.sharepoint.client.DriveItemPage
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Test
import org.mockito.kotlin.mock
import org.mockito.kotlin.verify
import org.mockito.kotlin.whenever

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
}
