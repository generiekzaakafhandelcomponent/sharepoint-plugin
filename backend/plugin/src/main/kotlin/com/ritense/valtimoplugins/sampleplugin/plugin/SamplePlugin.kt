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

import com.ritense.plugin.annotation.Plugin
import com.ritense.plugin.annotation.PluginAction
import com.ritense.plugin.annotation.PluginActionProperty
import com.ritense.plugin.annotation.PluginProperty
import com.ritense.processlink.domain.ActivityTypeWithEventName.SERVICE_TASK_START
import com.ritense.valtimoplugins.sampleplugin.client.SampleService
import io.github.oshai.kotlinlogging.KotlinLogging
import org.operaton.bpm.engine.delegate.DelegateExecution

private val logger = KotlinLogging.logger {}

/**
 * Sample plugin demonstrating a simple action that interacts with an API endpoint.
 * Note that the key in the @Plugin annotation must be unique, and
 * should be equal to the pluginId in the plugin's frontend configuration.
 */
@Plugin(
    key = "sample-plugin",
    title = "Sample Plugin",
    description = "This is a sample plugin demonstrating an API call action.",
)
open class SamplePlugin(
    private val sampleService: SampleService,
) {
    @PluginProperty(key = "apiUrl", secret = false)
    lateinit var apiUrl: String

    /**
     * Example action
     * Sends a GET request to an API endpoint and returns the response.
     */
    @PluginAction(
        key = "time-api-sample-action",
        title = "Time API test action",
        description = "Time API plugin action",
        activityTypes = [SERVICE_TASK_START],
    )
    open fun getCurrentTime(
        execution: DelegateExecution,
        @PluginActionProperty message: String,
    ): String {
        try {
            val result = sampleService.printAPIResults(apiUrl = apiUrl)
            logger.info { "Message: $message, Result: $result" }
            execution.setVariable("message", message)
            execution.setVariable("apiResult", result)
            return result
        } catch (e: Exception) {
            logger.error(e) { "Error: ${e.cause}" }
            return "Error: ${e.message}"
        }
    }
}
