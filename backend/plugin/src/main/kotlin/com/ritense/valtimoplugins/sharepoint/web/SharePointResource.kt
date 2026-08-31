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

import com.ritense.plugin.domain.PluginConfigurationId
import com.ritense.plugin.service.PluginService
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPlugin
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.valtimoplugins.sharepoint.domain.TestConnectionRequest
import com.ritense.valtimoplugins.sharepoint.domain.TestConnectionResult
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentPage
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
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
) {
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
}
