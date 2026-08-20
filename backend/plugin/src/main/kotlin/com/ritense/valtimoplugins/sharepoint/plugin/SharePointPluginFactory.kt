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

import com.fasterxml.jackson.databind.ObjectMapper
import com.ritense.plugin.PluginFactory
import com.ritense.plugin.service.PluginService
import com.ritense.processdocument.service.ProcessDocumentService
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sharepoint.service.SharePointService
import com.ritense.zakenapi.resolver.ZaakValueResolverFactory
import org.springframework.stereotype.Component

@SkipComponentScan
@Component
class SharePointPluginFactory(
    pluginService: PluginService,
    private val sharePointService: SharePointService,
    private val zaakValueResolverFactory: ZaakValueResolverFactory,
    private val processDocumentService: ProcessDocumentService,
) : PluginFactory<SharePointPlugin>(pluginService) {
    override fun create(): SharePointPlugin =
        SharePointPlugin(sharePointService, zaakValueResolverFactory, processDocumentService)
}
