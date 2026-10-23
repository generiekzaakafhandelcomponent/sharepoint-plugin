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

package com.ritense.valtimoplugins.sharepoint.client

import com.microsoft.graph.serviceclient.GraphServiceClient
import okhttp3.OkHttpClient
import okhttp3.mockwebserver.MockResponse
import okhttp3.mockwebserver.MockWebServer
import org.junit.jupiter.api.AfterEach
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.BeforeEach
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream

internal class MicrosoftGraphClientUploadTest {
    private val server = MockWebServer()
    private lateinit var client: MicrosoftGraphClient

    @BeforeEach
    fun setUp() {
        server.start()
        val graphServiceClient = GraphServiceClient(OkHttpClient())
        graphServiceClient.requestAdapter.baseUrl = server.url("/v1.0").toString().trimEnd('/')
        client = MicrosoftGraphClient(graphServiceClient)
    }

    @AfterEach
    fun tearDown() {
        server.shutdown()
    }

    @Test
    fun `uploadFile opens an upload session in the folder and sends the whole file through it`() {
        val content = "zaak document content".toByteArray()
        server.enqueue(
            MockResponse()
                .setHeader("Content-Type", "application/json")
                .setBody(
                    """{"uploadUrl":"${server.url("/upload-session/1")}","expirationDateTime":"2099-01-01T00:00:00Z",""" +
                        """"nextExpectedRanges":["0-"]}""",
                ),
        )
        server.enqueue(
            MockResponse()
                .setResponseCode(201)
                .setHeader("Content-Type", "application/json")
                .setBody("""{"id":"item-2","name":"brief 1.docx","webUrl":"https://sharepoint.example/brief 1.docx","file":{}}"""),
        )

        val uploaded = client.uploadFile("drive-1", "folder-1", "brief.docx", ByteArrayInputStream(content), content.size.toLong())

        val sessionRequest = server.takeRequest()
        assertEquals("POST", sessionRequest.method)
        assertEquals(
            "/v1.0/drives/drive-1/items/folder-1:/brief.docx:/createUploadSession",
            sessionRequest.requestUrl?.encodedPath?.replace("%3A", ":")?.replace("%2F", "/"),
        )
        assertTrue(sessionRequest.body.readUtf8().contains(""""@microsoft.graph.conflictBehavior":"rename""""))

        val uploadRequest = server.takeRequest()
        assertEquals("PUT", uploadRequest.method)
        assertEquals("/upload-session/1", uploadRequest.requestUrl?.encodedPath)
        assertEquals("bytes 0-${content.size - 1}/${content.size}", uploadRequest.getHeader("Content-Range"))
        assertEquals(String(content), uploadRequest.body.readUtf8())

        assertEquals("item-2", uploaded.id)
        assertEquals("brief 1.docx", uploaded.name)
    }
}
