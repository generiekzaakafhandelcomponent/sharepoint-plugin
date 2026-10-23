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

import com.fasterxml.jackson.core.JsonPointer
import com.microsoft.graph.models.odataerrors.ODataError
import com.ritense.document.exception.DocumentNotFoundException
import com.ritense.document.service.DocumentService
import com.ritense.plugin.domain.PluginConfigurationId
import com.ritense.plugin.service.PluginService
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPlugin
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.valtimoplugins.sharepoint.domain.StoreInDocumentenApiRequest
import com.ritense.valtimoplugins.sharepoint.domain.StoreInDocumentenApiResult
import com.ritense.valtimoplugins.sharepoint.domain.TestConnectionRequest
import com.ritense.valtimoplugins.sharepoint.domain.TestConnectionResult
import com.ritense.valtimoplugins.sharepoint.domain.UploadFromDocumentenApiRequest
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocument
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentNotFoundException
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentPage
import com.ritense.valtimoplugins.sharepoint.service.SharePointCaseService
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.DeleteMapping
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PathVariable
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.RestController
import org.springframework.web.server.ResponseStatusException
import java.util.UUID

@SkipComponentScan
@RestController
@RequestMapping("/api/v1/plugin/sharepoint")
class SharePointResource(

    private val pluginService: PluginService,
    private val sharePointService: SharePointService,
    private val documentService: DocumentService,
    private val sharePointCaseService: SharePointCaseService,
) {
    @Deprecated("Takes the folder from the browser; use the case-scoped documents/{documentId}/work-documents")
    @GetMapping("/{pluginConfigurationId}/work-documents")
    fun getWorkDocuments(
        @PathVariable pluginConfigurationId: UUID,
        @RequestParam docDefinition: String,
        @RequestParam year: String,
        @RequestParam zaaknummer: String,
        @RequestParam(required = false) pageSize: Int?,
        @RequestParam(required = false) nextLink: String?,
    ): ResponseEntity<WorkDocumentPage> {
        val configId = PluginConfigurationId.existingId(pluginConfigurationId)
        val plugin = pluginService.createInstance(configId) as? SharePointPlugin
            ?: throw ResponseStatusException(HttpStatus.NOT_FOUND, "Plugin configuration not found: $pluginConfigurationId")

        val page = try {
            sharePointService.listWorkDocuments(
                graphClient = plugin.microsoftGraphClient,
                driveId = plugin.driveId,
                dossierDefinitionName = docDefinition,
                year = year,
                zaaknummer = zaaknummer,
                pageSize = pageSize,
                nextLink = nextLink,
            )
        } catch (e: IllegalArgumentException) {
            throw ResponseStatusException(HttpStatus.BAD_REQUEST, e.message)
        }
        return ResponseEntity.ok(page)
    }

    @GetMapping("/{pluginConfigurationId}/documents/{documentId}/work-documents")
    fun getCaseWorkDocuments(
        @PathVariable pluginConfigurationId: UUID,
        @PathVariable documentId: UUID,
        @RequestParam(required = false) pageSize: Int?,
        @RequestParam(required = false) nextLink: String?,
    ): ResponseEntity<WorkDocumentPage> = handleErrors {
        ResponseEntity.ok(sharePointCaseService.listWorkDocuments(pluginConfigurationId, documentId, pageSize, nextLink))
    }

    @PostMapping("/{pluginConfigurationId}/documents/{documentId}/work-documents")
    fun uploadFromDocumentenApi(
        @PathVariable pluginConfigurationId: UUID,
        @PathVariable documentId: UUID,
        @RequestBody request: UploadFromDocumentenApiRequest,
    ): ResponseEntity<WorkDocument> = handleErrors {
        ResponseEntity.ok(sharePointCaseService.uploadFromDocumentenApi(pluginConfigurationId, documentId, request))
    }

    @PostMapping("/{pluginConfigurationId}/documents/{documentId}/work-documents/{itemId}/documenten-api")
    fun storeInDocumentenApi(
        @PathVariable pluginConfigurationId: UUID,
        @PathVariable documentId: UUID,
        @PathVariable itemId: String,
        @RequestBody request: StoreInDocumentenApiRequest,
    ): ResponseEntity<StoreInDocumentenApiResult> = handleErrors {
        ResponseEntity.ok(
            sharePointCaseService.storeInDocumentenApi(pluginConfigurationId, documentId, itemId, request.metadata)
        )
    }

    @DeleteMapping("/{pluginConfigurationId}/documents/{documentId}/work-documents/{itemId}")
    fun deleteWorkDocument(
        @PathVariable pluginConfigurationId: UUID,
        @PathVariable documentId: UUID,
        @PathVariable itemId: String,
    ): ResponseEntity<Unit> = handleErrors {
        sharePointCaseService.deleteWorkDocument(pluginConfigurationId, documentId, itemId)
        ResponseEntity.noContent().build()
    }

    @GetMapping("/documents/{documentId}/sharepoint-location")
    fun getSharePointLocation(
        @PathVariable documentId: String
    ): ResponseEntity<String> {
        val document = try {
            documentService.get(documentId)
        } catch (e: DocumentNotFoundException) {
            throw ResponseStatusException(HttpStatus.NOT_FOUND, "Document not found: $documentId")
        }
        val location = document.content().getValueBy(JsonPointer.compile("/sharePointMap"))
            .orElseThrow { ResponseStatusException(HttpStatus.NOT_FOUND, "No value found at path: sharePointMap") }
        return ResponseEntity.ok(location.asText())
    }

    @PostMapping("/test-connection")
    fun testConnection(@RequestBody request: TestConnectionRequest): ResponseEntity<TestConnectionResult> {
        val result = sharePointService.testConnection(
            tenantId = request.tenantId,
            clientId = request.clientId,
            clientSecret = request.clientSecret,
            hostname = request.hostname,
            sharePointSiteName = request.sharePointSiteName,
            baseFolderPath = request.baseFolderPath,
        )
        return ResponseEntity.ok(result)
    }

    private fun <T> handleErrors(block: () -> ResponseEntity<T>): ResponseEntity<T> {
        return try {
            block()
        } catch (e: DocumentNotFoundException) {
            throw ResponseStatusException(HttpStatus.NOT_FOUND, e.message)
        } catch (e: WorkDocumentNotFoundException) {
            throw ResponseStatusException(HttpStatus.NOT_FOUND, e.message)
        } catch (e: IllegalArgumentException) {
            throw ResponseStatusException(HttpStatus.BAD_REQUEST, e.message)
        } catch (e: IllegalStateException) {
            throw ResponseStatusException(HttpStatus.CONFLICT, e.message)
        } catch (e: ODataError) {
            // Graph's 401 and 403 concern the app registration, not the user's session
            val status = HttpStatus.resolve(e.responseStatusCode)
                ?.takeIf { it.is4xxClientError && it != HttpStatus.UNAUTHORIZED && it != HttpStatus.FORBIDDEN }
                ?: HttpStatus.BAD_GATEWAY
            throw ResponseStatusException(status, "SharePoint refused the request: ${e.error?.message ?: e.responseStatusCode}")
        }
    }
}
