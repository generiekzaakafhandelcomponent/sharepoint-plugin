package com.ritense.valtimoplugins.sharepoint.security

import com.ritense.valtimo.contract.authentication.AuthoritiesConstants.USER
import com.ritense.valtimo.contract.security.config.HttpConfigurerConfigurationException

import com.ritense.valtimo.contract.security.config.HttpSecurityConfigurer
import org.springframework.http.HttpMethod.GET
import org.springframework.security.config.annotation.web.builders.HttpSecurity


class SharePointHttpSecurityConfigurer: HttpSecurityConfigurer {
    override fun configure(http: HttpSecurity) {
        try {
            http.authorizeHttpRequests { requests ->
                requests
                    .requestMatchers(GET, "/api/v1/plugin/sharepoint/*/work-documents")
                    .hasAuthority(USER)
            }
        } catch (e: Exception) {
            throw HttpConfigurerConfigurationException(e)
        }
    }
}
