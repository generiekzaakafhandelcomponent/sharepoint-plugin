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
import org.springframework.stereotype.Component
import org.springframework.web.client.RestClient

@SkipComponentScan
@Component
class SampleClient(
    private val restClient: RestClient = RestClient.create(),
) {
    fun fetchTimeAPI(apiUrl: String): ApiResponse =
        try {
            val response =
                restClient
                    .get()
                    .uri(apiUrl)
                    .retrieve()
                    .toEntity(Timezone::class.java)

            val status = response.statusCode.value().toString()

            if (status != "200") {
                ApiResponse(error = "Error: status code $status")
            }

            ApiResponse(result = response, responseStatus = status)
        } catch (e: Exception) {
            ApiResponse(error = "Error: ${e.message}")
        }
}
