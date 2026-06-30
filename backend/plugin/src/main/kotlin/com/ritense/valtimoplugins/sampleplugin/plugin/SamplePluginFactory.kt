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

package com.ritense.valtimoplugins.sampleplugin.plugin

import com.ritense.plugin.PluginFactory
import com.ritense.plugin.service.PluginService
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimoplugins.sampleplugin.client.SampleService
import org.springframework.stereotype.Component

/**
 * Factory class to create instances of the SamplePlugin.
 * This is required for the plugin framework to instantiate the plugin.
 */
@SkipComponentScan
@Component
class SamplePluginFactory(
    pluginService: PluginService,
    val sampleService: SampleService,
) : PluginFactory<SamplePlugin>(pluginService) {
    override fun create(): SamplePlugin = SamplePlugin(sampleService)
}
