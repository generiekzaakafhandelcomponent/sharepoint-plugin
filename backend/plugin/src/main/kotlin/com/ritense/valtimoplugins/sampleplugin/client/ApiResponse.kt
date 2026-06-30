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

import org.springframework.http.ResponseEntity

data class Timezone(
    val year: Int? = null,
    val month: Int? = null,
    val day: Int? = null,
    val hour: Int? = null,
    val minute: Int? = null,
    val seconds: Int? = null,
    val milliSeconds: Int? = null,
    val dateTime: String? = null,
    val date: String? = null,
    val time: String? = null,
    val timeZone: String? = null,
    val dayOfWeek: String? = null,
    val dstActive: Boolean? = null,
)

data class ApiResponse(
    val result: ResponseEntity<Timezone?>? = null,
    val responseStatus: String? = null,
    val error: String? = null,
)
