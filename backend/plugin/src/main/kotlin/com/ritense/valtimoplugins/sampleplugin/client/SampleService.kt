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

package com.ritense.valtimoplugins.sampleplugin.client

import com.ritense.valtimo.contract.annotation.SkipComponentScan
import org.springframework.stereotype.Service

@SkipComponentScan
@Service
class SampleService(
    private val sampleClient: SampleClient,
) {
    fun printAPIResults(apiUrl: String): String {
        val apiResponse = sampleClient.fetchTimeAPI(apiUrl)

        if (apiResponse.error != null) {
            return "Failed: ${apiResponse.error}"
        }

        val tz = apiResponse.result?.body
        return "Timezone: ${tz?.timeZone}, DateTime: ${tz?.dateTime}, " +
            "DayOfWeek: ${tz?.dayOfWeek}, HTTP Status: ${apiResponse.responseStatus}"
    }
}
