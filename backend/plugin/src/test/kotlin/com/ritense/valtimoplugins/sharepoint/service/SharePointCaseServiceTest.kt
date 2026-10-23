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
import com.ritense.authorization.AuthorizationService
import com.ritense.document.domain.impl.JsonSchemaDocument
import com.ritense.document.service.DocumentService
import com.ritense.documentenapi.client.DocumentInformatieObject
import com.ritense.plugin.domain.PluginConfigurationId
import com.ritense.plugin.service.PluginService
import com.ritense.resource.service.TemporaryResourceStorageService
import com.ritense.temporaryresource.domain.StorageMetadataKeys
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.domain.StoreInDocumentenApiResult
import com.ritense.valtimoplugins.sharepoint.domain.UploadFromDocumentenApiRequest
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocument
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentPage
import com.ritense.valtimoplugins.sharepoint.domain.ZaakFolderLocation
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPlugin
import com.ritense.zakenapi.resolver.ZaakValueResolverFactory
import com.ritense.zakenapi.service.ZaakDocumentService
import com.ritense.zakenapi.uploadprocess.UploadProcessService
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.Test
import org.mockito.Answers
import org.mockito.kotlin.any
import org.mockito.kotlin.argumentCaptor
import org.mockito.kotlin.doThrow
import org.mockito.kotlin.eq
import org.mockito.kotlin.mock
import org.mockito.kotlin.never
import org.mockito.kotlin.verify
import org.mockito.kotlin.verifyNoInteractions
import org.mockito.kotlin.whenever
import org.springframework.context.ApplicationEventPublisher
import org.springframework.security.access.AccessDeniedException
import java.io.ByteArrayInputStream
import java.time.LocalDateTime
import java.util.UUID
import java.util.function.Function

internal class SharePointCaseServiceTest {
    private val pluginService: PluginService = mock()
    private val sharePointService: SharePointService = mock()
    private val documentService: DocumentService = mock()
    private val authorizationService: AuthorizationService = mock()
    private val zaakValueResolverFactory: ZaakValueResolverFactory = mock()
    private val zaakDocumentService: ZaakDocumentService = mock()
    private val temporaryResourceStorageService: TemporaryResourceStorageService = mock()
    private val uploadProcessService: UploadProcessService = mock()
    private val applicationEventPublisher: ApplicationEventPublisher = mock()
    private val graphClient: MicrosoftGraphClient = mock()
    private val plugin: SharePointPlugin = mock()
    private val document: JsonSchemaDocument = mock(defaultAnswer = Answers.RETURNS_DEEP_STUBS)

    private val caseService =
        SharePointCaseService(
            pluginService,
            sharePointService,
            documentService,
            authorizationService,
            zaakValueResolverFactory,
            zaakDocumentService,
            temporaryResourceStorageService,
            uploadProcessService,
            applicationEventPublisher,
        )

    private val pluginConfigurationId = UUID.randomUUID()
    private val documentId = UUID.randomUUID()
    private val location = ZaakFolderLocation("example", "2026", "ZAAK-001")
    private val sharePointFile =
        DriveItem().apply {
            id = "item-1"
            name = "brief.docx"
            file = File().apply { mimeType = "application/msword" }
        }

    @BeforeEach
    fun setUp() {
        whenever(documentService.get(documentId.toString())).thenReturn(document)
        whenever(document.definitionId().name()).thenReturn("example")
        whenever(document.createdOn()).thenReturn(LocalDateTime.of(2026, 3, 1, 12, 0))
        whenever(pluginService.createInstance(any<PluginConfigurationId>())).thenReturn(plugin)
        whenever(plugin.microsoftGraphClient).thenReturn(graphClient)
        whenever(plugin.driveId).thenReturn("drive-1")
        whenever(zaakValueResolverFactory.createResolver(documentId.toString())).thenReturn(Function { "ZAAK-001" })
    }

    @Test
    fun `listWorkDocuments lists the folder of the case's creation year and zaaknummer`() {
        val page = WorkDocumentPage(emptyList(), null)
        whenever(sharePointService.listZaakWorkDocuments(graphClient, "drive-1", location, 25, null)).thenReturn(page)

        assertEquals(page, caseService.listWorkDocuments(pluginConfigurationId, documentId, 25, null))
    }

    @Test
    fun `listWorkDocuments lists nothing for a user who may not view the case`() {
        whenever(documentService.get(documentId.toString())).thenThrow(AccessDeniedException("no"))

        assertThrows(AccessDeniedException::class.java) {
            caseService.listWorkDocuments(pluginConfigurationId, documentId, 25, null)
        }
        verifyNoInteractions(sharePointService)
    }

    @Test
    fun `uploadFromDocumentenApi uploads the zaak document's content under its bestandsnaam`() {
        val content = ByteArrayInputStream(ByteArray(3))
        val informatieObject: DocumentInformatieObject = mock()
        whenever(informatieObject.bestandsnaam).thenReturn("besluit.pdf")
        whenever(
            zaakDocumentService.getInformatieObject("documenten-config", documentId, "doc-1"),
        ).thenReturn(informatieObject)
        whenever(
            zaakDocumentService.downloadInformatieObject("documenten-config", documentId, "doc-1"),
        ).thenReturn(content)
        val uploaded =
            WorkDocument("item-2", "besluit.pdf", "https://sharepoint.example/besluit.pdf", 3, null, null, null)
        whenever(
            sharePointService.uploadToZaakFolder(graphClient, "drive-1", location, "besluit.pdf", content),
        ).thenReturn(uploaded)

        val result =
            caseService.uploadFromDocumentenApi(
                pluginConfigurationId,
                documentId,
                UploadFromDocumentenApiRequest("documenten-config", "doc-1"),
            )

        assertEquals(uploaded, result)
    }

    @Test
    fun `uploadFromDocumentenApi uses the titel when the zaak document has no bestandsnaam`() {
        val informatieObject: DocumentInformatieObject = mock()
        whenever(informatieObject.bestandsnaam).thenReturn(null)
        whenever(informatieObject.titel).thenReturn("Besluit")
        whenever(zaakDocumentService.getInformatieObject(any(), any(), any())).thenReturn(informatieObject)
        whenever(
            zaakDocumentService.downloadInformatieObject(any(), any(), any()),
        ).thenReturn(ByteArrayInputStream(ByteArray(0)))

        caseService.uploadFromDocumentenApi(
            pluginConfigurationId,
            documentId,
            UploadFromDocumentenApiRequest("c", "doc-1"),
        )

        verify(sharePointService).uploadToZaakFolder(eq(graphClient), eq("drive-1"), eq(location), eq("Besluit"), any())
    }

    @Test
    fun `uploadFromDocumentenApi writes nothing when the document is not related to the zaak`() {
        whenever(zaakDocumentService.getInformatieObject(any(), any(), any()))
            .thenThrow(IllegalArgumentException("InformatieObject is not related to this Zaak"))

        assertThrows(IllegalArgumentException::class.java) {
            caseService.uploadFromDocumentenApi(
                pluginConfigurationId,
                documentId,
                UploadFromDocumentenApiRequest("c", "doc-9"),
            )
        }
        verify(sharePointService, never()).uploadToZaakFolder(any(), any(), any(), any(), any())
    }

    @Test
    fun `uploadFromDocumentenApi writes nothing when the user may not view the zaak document`() {
        whenever(zaakDocumentService.getInformatieObject(any(), any(), any())).thenThrow(AccessDeniedException("no"))

        assertThrows(AccessDeniedException::class.java) {
            caseService.uploadFromDocumentenApi(
                pluginConfigurationId,
                documentId,
                UploadFromDocumentenApiRequest("c", "doc-1"),
            )
        }
        verify(sharePointService, never()).uploadToZaakFolder(any(), any(), any(), any(), any())
    }

    @Test
    fun `uploadFromDocumentenApi needs MODIFY on the case`() {
        denyModify()

        assertThrows(AccessDeniedException::class.java) {
            caseService.uploadFromDocumentenApi(
                pluginConfigurationId,
                documentId,
                UploadFromDocumentenApiRequest("c", "doc-1"),
            )
        }
        verifyNoInteractions(zaakDocumentService)
        verify(sharePointService, never()).uploadToZaakFolder(any(), any(), any(), any(), any())
    }

    @Test
    fun `storeInDocumentenApi stores the file with the dialog's metadata, then removes it from SharePoint`() {
        givenStoredAs("https://documenten.example/enkelvoudiginformatieobjecten/1")

        val result =
            caseService.storeInDocumentenApi(
                pluginConfigurationId,
                documentId,
                "item-1",
                mapOf(
                    "titel" to "Brief",
                    "informatieobjecttype" to "https://type",
                    "beschrijving" to "",
                ),
            )

        assertEquals(StoreInDocumentenApiResult(stored = true, removedFromSharePoint = true), result)
        val metadata = argumentCaptor<Map<String, Any>>()
        verify(temporaryResourceStorageService).store(any(), metadata.capture())
        assertEquals("Brief", metadata.firstValue["titel"])
        assertEquals("https://type", metadata.firstValue["informatieobjecttype"])
        assertEquals("brief.docx", metadata.firstValue["filename"])
        assertEquals("application/msword", metadata.firstValue["contentType"])
        assertEquals(false, metadata.firstValue.containsKey("beschrijving"))
        verify(uploadProcessService).startUploadResourceProcess(documentId, "resource-1")
        verify(sharePointService).deleteZaakWorkDocument(graphClient, "drive-1", location, "item-1")
    }

    @Test
    fun `storeInDocumentenApi passes only the dialog's fields and sets the file details itself`() {
        givenStoredAs("https://documenten.example/enkelvoudiginformatieobjecten/1")

        caseService.storeInDocumentenApi(
            pluginConfigurationId,
            documentId,
            "item-1",
            mapOf(
                "titel" to "Brief",
                "documentUrl" to "https://elsewhere",
                "filename" to "other.exe",
                "processInstanceId" to "p",
            ),
        )

        val metadata = argumentCaptor<Map<String, Any>>()
        verify(temporaryResourceStorageService).store(any(), metadata.capture())
        assertEquals(
            mapOf("titel" to "Brief", "filename" to "brief.docx", "contentType" to "application/msword"),
            metadata.firstValue,
        )
    }

    @Test
    fun `storeInDocumentenApi keeps the SharePoint file while the document is not stored yet`() {
        givenStoredAs(null)

        val result =
            caseService.storeInDocumentenApi(
                pluginConfigurationId,
                documentId,
                "item-1",
                mapOf(
                    "titel" to "Brief",
                ),
            )

        assertEquals(StoreInDocumentenApiResult(stored = false, removedFromSharePoint = false), result)
        verify(sharePointService, never()).deleteZaakWorkDocument(any(), any(), any(), any())
    }

    @Test
    fun `storeInDocumentenApi reports a stored document that SharePoint would not remove`() {
        givenStoredAs("https://documenten.example/enkelvoudiginformatieobjecten/1")
        whenever(sharePointService.deleteZaakWorkDocument(graphClient, "drive-1", location, "item-1"))
            .thenThrow(IllegalStateException("locked"))

        val result =
            caseService.storeInDocumentenApi(
                pluginConfigurationId,
                documentId,
                "item-1",
                mapOf(
                    "titel" to "Brief",
                ),
            )

        assertEquals(StoreInDocumentenApiResult(stored = true, removedFromSharePoint = false), result)
    }

    @Test
    fun `storeInDocumentenApi keeps the SharePoint file when the upload process fails`() {
        givenStoredAs(null)
        whenever(uploadProcessService.startUploadResourceProcess(documentId, "resource-1"))
            .thenThrow(IllegalStateException("No upload-process linked to case"))

        assertThrows(IllegalStateException::class.java) {
            caseService.storeInDocumentenApi(pluginConfigurationId, documentId, "item-1", mapOf("titel" to "Brief"))
        }
        verify(sharePointService, never()).deleteZaakWorkDocument(any(), any(), any(), any())
    }

    @Test
    fun `storeInDocumentenApi starts no upload process for a SharePoint folder`() {
        whenever(sharePointService.getZaakWorkDocument(graphClient, "drive-1", location, "sub"))
            .thenThrow(IllegalArgumentException("not a file"))

        assertThrows(IllegalArgumentException::class.java) {
            caseService.storeInDocumentenApi(pluginConfigurationId, documentId, "sub", mapOf("titel" to "Map"))
        }
        verifyNoInteractions(uploadProcessService, temporaryResourceStorageService)
    }

    @Test
    fun `storeInDocumentenApi needs MODIFY on the case`() {
        denyModify()

        assertThrows(AccessDeniedException::class.java) {
            caseService.storeInDocumentenApi(pluginConfigurationId, documentId, "item-1", mapOf("titel" to "Brief"))
        }
        verifyNoInteractions(uploadProcessService, temporaryResourceStorageService)
        verify(sharePointService, never()).deleteZaakWorkDocument(any(), any(), any(), any())
    }

    @Test
    fun `deleteWorkDocument deletes the item from the case's zaak folder`() {
        caseService.deleteWorkDocument(pluginConfigurationId, documentId, "item-1")

        verify(sharePointService).deleteZaakWorkDocument(graphClient, "drive-1", location, "item-1")
    }

    @Test
    fun `deleteWorkDocument needs MODIFY on the case`() {
        denyModify()

        assertThrows(AccessDeniedException::class.java) {
            caseService.deleteWorkDocument(pluginConfigurationId, documentId, "item-1")
        }
        verify(sharePointService, never()).deleteZaakWorkDocument(any(), any(), any(), any())
    }

    private fun denyModify() {
        doThrow(AccessDeniedException("no")).whenever(authorizationService).requirePermission<JsonSchemaDocument>(any())
    }

    private fun givenStoredAs(documentUrl: String?) {
        whenever(
            sharePointService.getZaakWorkDocument(graphClient, "drive-1", location, "item-1"),
        ).thenReturn(sharePointFile)
        whenever(
            sharePointService.downloadContent(graphClient, "drive-1", sharePointFile),
        ).thenReturn(ByteArrayInputStream(ByteArray(4)))
        whenever(temporaryResourceStorageService.store(any(), any())).thenReturn("resource-1")
        whenever(temporaryResourceStorageService.getMetadataValueOrNull("resource-1", StorageMetadataKeys.DOCUMENT_URL))
            .thenReturn(documentUrl)
    }
}
