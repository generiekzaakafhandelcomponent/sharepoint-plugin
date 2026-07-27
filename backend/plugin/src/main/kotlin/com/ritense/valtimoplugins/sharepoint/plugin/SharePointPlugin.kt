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

import com.azure.identity.ClientSecretCredentialBuilder
import com.fasterxml.jackson.databind.ObjectMapper
import com.microsoft.graph.serviceclient.GraphServiceClient
import com.ritense.plugin.annotation.Plugin
import com.ritense.plugin.annotation.PluginAction
import com.ritense.plugin.annotation.PluginActionProperty
import com.ritense.plugin.annotation.PluginEvent
import com.ritense.plugin.annotation.PluginProperty
import com.ritense.plugin.domain.EventType
import com.ritense.processlink.domain.ActivityTypeWithEventName.SERVICE_TASK_START
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import io.github.oshai.kotlinlogging.KotlinLogging
import org.operaton.bpm.engine.delegate.DelegateExecution

private val logger = KotlinLogging.logger {}

@Plugin(
    key = "sharepoint-plugin",
    title = "SharePoint Plugin",
    description = "Plugin voor integratie met Microsoft SharePoint voor werkdocumentenbeheer.",
)
open class SharePointPlugin(
    private val sharePointService: SharePointService,
    private val objectMapper: ObjectMapper,
) {
    @PluginProperty(key = "tenantId", secret = false)
    lateinit var tenantId: String

    @PluginProperty(key = "clientId", secret = false)
    lateinit var clientId: String

    @PluginProperty(key = "clientSecret", secret = true)
    lateinit var clientSecret: String

    @PluginProperty(key = "hostname", secret = false)
    lateinit var hostname: String

    @PluginProperty(key = "sharePointSiteName", secret = false)
    lateinit var sharePointSiteName: String

    @PluginProperty(key = "baseFolderPath", secret = false)
    lateinit var baseFolderPath: String

    var driveId: String = "init"

    val microsoftGraphClient: MicrosoftGraphClient by lazy {
        val credential = ClientSecretCredentialBuilder()
            .tenantId(tenantId)
            .clientId(clientId)
            .clientSecret(clientSecret)
            .build()
        MicrosoftGraphClient(GraphServiceClient(credential, "https://graph.microsoft.com/.default"))
    }

    @PluginEvent(invokedOn = [EventType.CREATE, EventType.UPDATE])
    fun getConfigProps() {
        try {
            val siteId = microsoftGraphClient.getSiteId(hostname, sharePointSiteName)
            driveId = microsoftGraphClient.getDriveIdByName(siteId, driveId)
        } catch (e: Exception) {
            logger.error(e) { "Failed to resolve SharePoint site/drive configuration for $hostname/sites/$sharePointSiteName" }
        }
    }

    @PluginAction(
        key = "create-zaak-folder",
        title = "Maak zaakmap aan in SharePoint",
        description = "Maakt een map aan in SharePoint voor de werkdocumenten van een zaak.",
        activityTypes = [SERVICE_TASK_START],
    )
    open fun createZaakFolder(
        execution: DelegateExecution,
        @PluginActionProperty zaaktypeVariable: String,
        @PluginActionProperty yearVariable: String,
        @PluginActionProperty zaaknummerVariable: String,
    ) {
        val zaaktype = execution.getVariable(zaaktypeVariable)?.toString()
            ?: error("Procesvariabele '$zaaktypeVariable' niet gevonden in executie ${execution.id}")
        val year = execution.getVariable(yearVariable)?.toString()
            ?: error("Procesvariabele '$yearVariable' niet gevonden in executie ${execution.id}")
        val zaaknummer = execution.getVariable(zaaknummerVariable)?.toString()
            ?: error("Procesvariabele '$zaaknummerVariable' niet gevonden in executie ${execution.id}")

        logger.info { "Creating SharePoint folder for zaak: $baseFolderPath/$zaaktype/$year/$zaaknummer" }
        sharePointService.createZaakFolder(
            graphClient = microsoftGraphClient,
            driveId = driveId,
            baseFolderPath = baseFolderPath,
            zaaktype = zaaktype,
            year = year,
            zaaknummer = zaaknummer,
        )
    }

    @PluginAction(
        key = "list-work-documents",
        title = "Haal werkdocumenten op uit SharePoint",
        description = "Haalt alle werkdocumenten op voor een zaak en slaat ze op als procesvariabele.",
        activityTypes = [SERVICE_TASK_START],
    )
    open fun listWorkDocuments(
        execution: DelegateExecution,
        @PluginActionProperty zaaktypeVariable: String,
        @PluginActionProperty yearVariable: String,
        @PluginActionProperty zaaknummerVariable: String,
        @PluginActionProperty resultVariable: String,
    ) {
        val zaaktype = execution.getVariable(zaaktypeVariable)?.toString()
            ?: error("Procesvariabele '$zaaktypeVariable' niet gevonden in executie ${execution.id}")
        val year = execution.getVariable(yearVariable)?.toString()
            ?: error("Procesvariabele '$yearVariable' niet gevonden in executie ${execution.id}")
        val zaaknummer = execution.getVariable(zaaknummerVariable)?.toString()
            ?: error("Procesvariabele '$zaaknummerVariable' niet gevonden in executie ${execution.id}")

        logger.info { "Listing SharePoint work documents for zaak: $baseFolderPath/$zaaktype/$year/$zaaknummer" }
        val documents = sharePointService.listWorkDocuments(
            graphClient = microsoftGraphClient,
            driveId = driveId,
            baseFolderPath = baseFolderPath,
            zaaktype = zaaktype,
            year = year,
            zaaknummer = zaaknummer,
        )

        execution.setVariable(resultVariable, objectMapper.writeValueAsString(documents))
    }
}
