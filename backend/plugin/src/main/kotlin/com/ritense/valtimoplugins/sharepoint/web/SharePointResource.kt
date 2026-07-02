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
import com.ritense.valtimoplugins.sharepoint.service.model.WorkDocument
import org.springframework.http.HttpStatus
import org.springframework.http.ResponseEntity
import org.springframework.stereotype.Component
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PathVariable
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
        @RequestParam zaaktype: String,
        @RequestParam year: String,
        @RequestParam zaaknummer: String,
    ): ResponseEntity<List<WorkDocument>> {
        val configId = PluginConfigurationId.existingId(pluginConfigurationId)
        val plugin = pluginService.createInstance(configId) as? SharePointPlugin
            ?: throw ResponseStatusException(HttpStatus.NOT_FOUND, "Plugin configuration not found: $pluginConfigurationId")

        val documents = sharePointService.listWorkDocuments(
            graphClient = plugin.microsoftGraphClient,
            driveId = plugin.driveId,
            baseFolderPath = plugin.baseFolderPath,
            zaaktype = zaaktype,
            year = year,
            zaaknummer = zaaknummer,
        )
        return ResponseEntity.ok(documents)
    }
}
