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

package com.ritense.valtimoplugins.sharepoint.autoconfiguration

import com.azure.identity.ClientSecretCredentialBuilder
import com.fasterxml.jackson.databind.ObjectMapper
import com.microsoft.graph.serviceclient.GraphServiceClient
import com.ritense.plugin.service.PluginService
import com.ritense.valtimoplugins.sharepoint.client.MicrosoftGraphClient
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPluginFactory
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.valtimoplugins.sharepoint.web.SharePointResource
import org.springframework.boot.autoconfigure.AutoConfiguration
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean
import org.springframework.boot.context.properties.EnableConfigurationProperties
import org.springframework.context.annotation.Bean

@AutoConfiguration
@EnableConfigurationProperties(MicrosoftGraphProperties::class)
class SharePointAutoConfiguration {
    @Bean
    @ConditionalOnMissingBean(GraphServiceClient::class)
    fun graphServiceClient(properties: MicrosoftGraphProperties): GraphServiceClient {
        val credential = ClientSecretCredentialBuilder()
            .tenantId(properties.tenantId)
            .clientId(properties.clientId)
            .clientSecret(properties.clientSecret)
            .build()
        return GraphServiceClient(credential, "https://graph.microsoft.com/.default")
    }

    @Bean
    @ConditionalOnMissingBean(MicrosoftGraphClient::class)
    fun microsoftGraphClient(graphServiceClient: GraphServiceClient): MicrosoftGraphClient =
        MicrosoftGraphClient(graphServiceClient)

    @Bean
    @ConditionalOnMissingBean(SharePointService::class)
    fun sharePointService(graphClient: MicrosoftGraphClient): SharePointService =
        SharePointService(graphClient)

    @Bean
    @ConditionalOnMissingBean(SharePointPluginFactory::class)
    fun sharePointPluginFactory(
        pluginService: PluginService,
        sharePointService: SharePointService,
        objectMapper: ObjectMapper,
    ): SharePointPluginFactory = SharePointPluginFactory(pluginService, sharePointService, objectMapper)

    @Bean
    @ConditionalOnMissingBean(SharePointResource::class)
    fun sharePointResource(
        pluginService: PluginService,
        sharePointService: SharePointService,
    ): SharePointResource = SharePointResource(pluginService, sharePointService)
}
