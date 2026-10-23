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

import com.ritense.authorization.AuthorizationService
import com.ritense.document.service.DocumentService
import com.ritense.plugin.service.PluginService
import com.ritense.resource.service.TemporaryResourceStorageService
import com.ritense.processdocument.service.ProcessDocumentService
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPluginFactory
import com.ritense.valtimoplugins.sharepoint.security.SharePointHttpSecurityConfigurer
import com.ritense.valtimoplugins.sharepoint.service.SharePointCaseService
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.valtimoplugins.sharepoint.web.SharePointResource
import com.ritense.zakenapi.resolver.ZaakValueResolverFactory
import com.ritense.zakenapi.service.ZaakDocumentService
import com.ritense.zakenapi.uploadprocess.UploadProcessService
import org.springframework.boot.autoconfigure.AutoConfiguration
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean
import org.springframework.context.ApplicationEventPublisher
import org.springframework.context.annotation.Bean
import org.springframework.core.annotation.Order

@AutoConfiguration
class SharePointAutoConfiguration {
    @Bean
    @ConditionalOnMissingBean(SharePointService::class)
    fun sharePointService(): SharePointService = SharePointService()

    @Bean
    @ConditionalOnMissingBean(SharePointPluginFactory::class)
    fun sharePointPluginFactory(
        pluginService: PluginService,
        sharePointService: SharePointService,
        zaakValueResolverFactory: ZaakValueResolverFactory,
        processDocumentService: ProcessDocumentService,
    ): SharePointPluginFactory =
        SharePointPluginFactory(pluginService, sharePointService, zaakValueResolverFactory, processDocumentService)

    @Bean
    @ConditionalOnMissingBean(SharePointCaseService::class)
    fun sharePointCaseService(
        pluginService: PluginService,
        sharePointService: SharePointService,
        documentService: DocumentService,
        authorizationService: AuthorizationService,
        zaakValueResolverFactory: ZaakValueResolverFactory,
        zaakDocumentService: ZaakDocumentService,
        temporaryResourceStorageService: TemporaryResourceStorageService,
        uploadProcessService: UploadProcessService,
        applicationEventPublisher: ApplicationEventPublisher,
    ): SharePointCaseService = SharePointCaseService(
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

    @Bean
    @ConditionalOnMissingBean(SharePointResource::class)
    fun sharePointResource(
        pluginService: PluginService,
        sharePointService: SharePointService,
        documentService: DocumentService,
        sharePointCaseService: SharePointCaseService,
    ): SharePointResource = SharePointResource(pluginService, sharePointService, documentService, sharePointCaseService)

    @Order(301)
    @Bean
    @ConditionalOnMissingBean(SharePointHttpSecurityConfigurer::class)
    fun sharepointHttpSecurityConfigurer(): SharePointHttpSecurityConfigurer = SharePointHttpSecurityConfigurer()
}
