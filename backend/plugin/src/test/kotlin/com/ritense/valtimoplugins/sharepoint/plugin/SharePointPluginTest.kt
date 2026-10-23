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

package com.ritense.valtimoplugins.sharepoint.plugin

import com.ritense.document.domain.Document
import com.ritense.processdocument.domain.impl.OperatonProcessInstanceId
import com.ritense.processdocument.service.ProcessDocumentService
import com.ritense.valtimoplugins.sharepoint.domain.ZaakFolderLocation
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.zakenapi.resolver.ZaakValueResolverFactory
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.Test
import org.mockito.Answers
import org.mockito.kotlin.any
import org.mockito.kotlin.eq
import org.mockito.kotlin.mock
import org.mockito.kotlin.verify
import org.mockito.kotlin.whenever
import org.operaton.bpm.engine.delegate.DelegateExecution
import java.time.LocalDateTime
import java.util.function.Function

internal class SharePointPluginTest {
    private val sharePointService: SharePointService = mock()
    private val zaakValueResolverFactory: ZaakValueResolverFactory = mock()
    private val processDocumentService: ProcessDocumentService = mock()
    private val execution: DelegateExecution = mock()
    private val document: Document = mock(defaultAnswer = Answers.RETURNS_DEEP_STUBS)

    private val plugin =
        SharePointPlugin(sharePointService, zaakValueResolverFactory, processDocumentService).apply {
            tenantId = "tenant-1"
            clientId = "client-1"
            clientSecret = "secret"
            hostname = "contoso.sharepoint.com"
            sharePointSiteName = "site"
            baseFolderPath = "zaakdossiers"
        }

    @BeforeEach
    fun setUp() {
        whenever(execution.businessKey).thenReturn("case-1")
        whenever(execution.processInstanceId).thenReturn("7c1c1f6a-1f39-4bd0-9b0e-2a1d6c1c9e01")
        whenever(zaakValueResolverFactory.createResolver("case-1")).thenReturn(Function { "ZAAK-001" })
        whenever(
            processDocumentService.getDocument(any<OperatonProcessInstanceId>(), eq(execution)),
        ).thenReturn(document)
        whenever(document.definitionId().name()).thenReturn("example")
        whenever(document.createdOn()).thenReturn(LocalDateTime.of(2026, 2, 1, 9, 0))
        whenever(sharePointService.resolveDriveId(any(), any(), any(), any())).thenReturn("drive-1")
    }

    @Test
    fun `create-zaak-folder creates the case's folder and stores its location`() {
        whenever(sharePointService.createZaakFolder(any(), eq("drive-1"), eq("example"), eq("2026"), eq("ZAAK-001")))
            .thenReturn("https://sharepoint.example/ZAAK-001")

        plugin.createZaakFolder(execution, "sharePointFolderLocation")

        verify(execution).setVariable("sharePointFolderLocation", "https://sharepoint.example/ZAAK-001")
    }

    @Test
    fun `delete-work-document deletes the named file from the case's zaak folder`() {
        plugin.deleteWorkDocument(execution, "brief.docx")

        verify(sharePointService).deleteZaakWorkDocumentByName(
            any(),
            eq("drive-1"),
            eq(ZaakFolderLocation("example", "2026", "ZAAK-001")),
            eq("brief.docx"),
        )
    }
}
