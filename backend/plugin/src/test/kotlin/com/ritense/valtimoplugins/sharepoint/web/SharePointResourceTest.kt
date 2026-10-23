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

package com.ritense.valtimoplugins.sharepoint.web

import com.microsoft.graph.models.odataerrors.ODataError
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentNotFoundException
import com.ritense.valtimoplugins.sharepoint.service.SharePointCaseService
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Test
import org.mockito.kotlin.doThrow
import org.mockito.kotlin.mock
import org.mockito.kotlin.whenever
import org.springframework.http.HttpStatus
import org.springframework.web.server.ResponseStatusException
import java.util.UUID

private class StatusODataError(
    statusCode: Int,
) : ODataError() {
    init {
        responseStatusCode = statusCode
    }
}

internal class SharePointResourceTest {
    private val sharePointCaseService: SharePointCaseService = mock()
    private val resource = SharePointResource(mock(), mock(), mock(), sharePointCaseService)
    private val pluginConfigurationId = UUID.randomUUID()
    private val documentId = UUID.randomUUID()

    @Test
    fun `delete reports a file SharePoint keeps locked as locked`() {
        doThrow(StatusODataError(423))
            .whenever(sharePointCaseService)
            .deleteWorkDocument(pluginConfigurationId, documentId, "item-1")

        val exception =
            assertThrows(ResponseStatusException::class.java) {
                resource.deleteWorkDocument(pluginConfigurationId, documentId, "item-1")
            }
        assertEquals(HttpStatus.LOCKED, exception.statusCode)
    }

    @Test
    fun `delete reports a refused app registration as a gateway error, not as the user's session`() {
        doThrow(StatusODataError(401))
            .whenever(sharePointCaseService)
            .deleteWorkDocument(pluginConfigurationId, documentId, "item-1")

        val exception =
            assertThrows(ResponseStatusException::class.java) {
                resource.deleteWorkDocument(pluginConfigurationId, documentId, "item-1")
            }
        assertEquals(HttpStatus.BAD_GATEWAY, exception.statusCode)
    }

    @Test
    fun `delete reports an item outside the zaak folder as not found`() {
        doThrow(WorkDocumentNotFoundException("not here"))
            .whenever(sharePointCaseService)
            .deleteWorkDocument(pluginConfigurationId, documentId, "item-9")

        val exception =
            assertThrows(ResponseStatusException::class.java) {
                resource.deleteWorkDocument(pluginConfigurationId, documentId, "item-9")
            }
        assertEquals(HttpStatus.NOT_FOUND, exception.statusCode)
    }

    @Test
    fun `delete answers no content once the file is deleted`() {
        assertEquals(
            HttpStatus.NO_CONTENT,
            resource.deleteWorkDocument(pluginConfigurationId, documentId, "item-1").statusCode,
        )
    }
}
