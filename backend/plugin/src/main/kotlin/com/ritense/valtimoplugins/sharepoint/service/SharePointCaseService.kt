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

package com.ritense.valtimoplugins.sharepoint.service

import com.ritense.authorization.AuthorizationService
import com.ritense.authorization.request.EntityAuthorizationRequest
import com.ritense.document.domain.impl.JsonSchemaDocument
import com.ritense.document.service.DocumentService
import com.ritense.document.service.JsonSchemaDocumentActionProvider.MODIFY
import com.ritense.documentenapi.domain.DocumentenApiUploadFieldKey
import com.ritense.plugin.domain.PluginConfigurationId
import com.ritense.plugin.service.PluginService
import com.ritense.resource.domain.MetadataType
import com.ritense.resource.domain.TemporaryResourceUploadedEvent
import com.ritense.resource.service.TemporaryResourceStorageService
import com.ritense.temporaryresource.domain.StorageMetadataKeys
import com.ritense.valtimo.contract.annotation.SkipComponentScan
import com.ritense.valtimo.contract.utils.SecurityUtils
import com.ritense.valtimoplugins.sharepoint.domain.StoreInDocumentenApiResult
import com.ritense.valtimoplugins.sharepoint.domain.UploadFromDocumentenApiRequest
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocument
import com.ritense.valtimoplugins.sharepoint.domain.WorkDocumentPage
import com.ritense.valtimoplugins.sharepoint.domain.ZaakFolderLocation
import com.ritense.valtimoplugins.sharepoint.plugin.SharePointPlugin
import com.ritense.zakenapi.resolver.ZaakValueResolverFactory
import com.ritense.zakenapi.service.ZaakDocumentService
import com.ritense.zakenapi.uploadprocess.UploadProcessService
import io.github.oshai.kotlinlogging.KotlinLogging
import org.springframework.context.ApplicationEventPublisher
import org.springframework.stereotype.Service
import java.util.UUID

private val logger = KotlinLogging.logger {}

@SkipComponentScan
@Service
class SharePointCaseService(
    private val pluginService: PluginService,
    private val sharePointService: SharePointService,
    private val documentService: DocumentService,
    private val authorizationService: AuthorizationService,
    private val zaakValueResolverFactory: ZaakValueResolverFactory,
    private val zaakDocumentService: ZaakDocumentService,
    private val temporaryResourceStorageService: TemporaryResourceStorageService,
    private val uploadProcessService: UploadProcessService,
    private val applicationEventPublisher: ApplicationEventPublisher,
) {
    fun listWorkDocuments(
        pluginConfigurationId: UUID,
        documentId: UUID,
        pageSize: Int?,
        nextLink: String?,
    ): WorkDocumentPage {
        val case = resolveCase(pluginConfigurationId, documentId, requireModify = false)
        return sharePointService.listZaakWorkDocuments(
            case.plugin.microsoftGraphClient,
            case.plugin.driveId,
            case.folder,
            pageSize,
            nextLink,
        )
    }

    fun uploadFromDocumentenApi(
        pluginConfigurationId: UUID,
        documentId: UUID,
        request: UploadFromDocumentenApiRequest,
    ): WorkDocument {
        val case = resolveCase(pluginConfigurationId, documentId, requireModify = true)
        val informatieObject =
            zaakDocumentService.getInformatieObject(
                request.documentenApiPluginConfigurationId,
                documentId,
                request.documentenApiDocumentId,
            )
        val content =
            zaakDocumentService.downloadInformatieObject(
                request.documentenApiPluginConfigurationId,
                documentId,
                request.documentenApiDocumentId,
            )
        return sharePointService.uploadToZaakFolder(
            case.plugin.microsoftGraphClient,
            case.plugin.driveId,
            case.folder,
            informatieObject.bestandsnaam?.takeIf { it.isNotBlank() } ?: informatieObject.titel,
            content,
        )
    }

    fun storeInDocumentenApi(
        pluginConfigurationId: UUID,
        documentId: UUID,
        itemId: String,
        metadata: Map<String, String>,
    ): StoreInDocumentenApiResult {
        val case = resolveCase(pluginConfigurationId, documentId, requireModify = true)
        val graphClient = case.plugin.microsoftGraphClient
        val item = sharePointService.getZaakWorkDocument(graphClient, case.plugin.driveId, case.folder, itemId)

        // Only the dialog's fields: a key like documentUrl would make core skip the upload process
        val resourceMetadata: MutableMap<String, Any> =
            metadata.filter { (key, value) -> key in UPLOAD_METADATA_FIELDS && value.isNotBlank() }.toMutableMap()
        item.name?.let { resourceMetadata[MetadataType.FILE_NAME.key] = it }
        item.file?.mimeType?.let { resourceMetadata[MetadataType.CONTENT_TYPE.key] = it }
        SecurityUtils.getCurrentUserLogin()?.let { resourceMetadata[MetadataType.USER.key] = it }

        val resourceId =
            sharePointService.downloadContent(graphClient, case.plugin.driveId, item).use {
                temporaryResourceStorageService.store(it, resourceMetadata)
            }
        applicationEventPublisher.publishEvent(TemporaryResourceUploadedEvent(resourceId))
        uploadProcessService.startUploadResourceProcess(documentId, resourceId)

        if (temporaryResourceStorageService.getMetadataValueOrNull(resourceId, StorageMetadataKeys.DOCUMENT_URL) ==
            null
        ) {
            logger.info {
                "SharePoint document '${item.name}' is not stored in the Documenten API yet; keeping it in SharePoint"
            }
            return StoreInDocumentenApiResult(stored = false, removedFromSharePoint = false)
        }
        val removed =
            try {
                sharePointService.deleteZaakWorkDocument(graphClient, case.plugin.driveId, case.folder, itemId)
                true
            } catch (e: Exception) {
                logger.warn(e) {
                    "Stored SharePoint document '${item.name}' in the Documenten API but could not remove it from SharePoint"
                }
                false
            }
        return StoreInDocumentenApiResult(stored = true, removedFromSharePoint = removed)
    }

    fun deleteWorkDocument(
        pluginConfigurationId: UUID,
        documentId: UUID,
        itemId: String,
    ) {
        val case = resolveCase(pluginConfigurationId, documentId, requireModify = true)
        sharePointService.deleteZaakWorkDocument(
            case.plugin.microsoftGraphClient,
            case.plugin.driveId,
            case.folder,
            itemId,
        )
    }

    private fun resolveCase(
        pluginConfigurationId: UUID,
        documentId: UUID,
        requireModify: Boolean,
    ): ResolvedCase {
        val document = documentService.get(documentId.toString())
        if (requireModify) {
            authorizationService.requirePermission(
                EntityAuthorizationRequest(JsonSchemaDocument::class.java, MODIFY, document as JsonSchemaDocument),
            )
        }
        val plugin =
            pluginService.createInstance(PluginConfigurationId.existingId(pluginConfigurationId)) as? SharePointPlugin
                ?: throw IllegalArgumentException("No SharePoint plugin configuration with id $pluginConfigurationId")
        val zaaknummer =
            zaakValueResolverFactory.createResolver(documentId.toString()).apply("identificatie") as? String
                ?: throw IllegalStateException("Case $documentId is not linked to a zaak")
        return ResolvedCase(plugin, ZaakFolderLocation.of(document, zaaknummer))
    }

    private data class ResolvedCase(
        val plugin: SharePointPlugin,
        val folder: ZaakFolderLocation,
    )

    companion object {
        private val UPLOAD_METADATA_FIELDS =
            DocumentenApiUploadFieldKey.entries.map { it.property }.toSet() + setOf("ontvangstdatum", "verzenddatum")
    }
}
