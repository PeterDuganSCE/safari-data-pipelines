import json
import logging
import time
import getpass

import msal
import requests


class GraphAuthenticationError(Exception):
    """Raised when Graph authentication fails."""
    pass


class GraphApiError(Exception):
    """Raised when a Graph API call fails."""
    pass


class GraphClient:

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

    def __init__(
        self,
        app_config_path="config/sharepoint_auth.json",
        credential_file=None,
        timeout=60
    ):

        self.app_config_path = app_config_path

        self.credential_file = (
            credential_file
            or f"C:/Users/{getpass.getuser()}/sharepoint_auth.json"  # Default location for the credential file
        )

        self.timeout = timeout

        self.session = requests.Session()

        self.headers = None
        self.app = None

        self.logger = logging.getLogger(__name__)

    ###########################################################################
    # Authentication
    ###########################################################################

    def authenticate(self):
        """
        Authenticate using ROPC and build headers.
        """

        self.logger.info("Authenticating with Microsoft Graph")

        with open(self.app_config_path, "r") as f:
            app_details = json.load(f)["graphAuth"]

        with open(self.credential_file, "r") as f:
            creds = json.load(f)["windowsAuth"]

        tenant_id = app_details["tenant_id"]
        client_id = app_details["client_id"]

        username = creds["email"]
        password = creds["password"]

        authority = (
            f"https://login.microsoftonline.com/{tenant_id}"
        )

        self.app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority
        )

        result = self.app.acquire_token_by_username_password(
            username=username,
            password=password,
            scopes=self.GRAPH_SCOPE
        )

        if not result or "access_token" not in result:

            raise GraphAuthenticationError(
                f"Authentication failed:\n"
                f"Error: {result.get('error')}\n"
                f"Description: {result.get('error_description')}"
            )

        self.headers = {
            "Authorization": (
                f"{result['token_type']} "
                f"{result['access_token']}"
            ),
            "Content-Type": "application/json"
        }

        me = self.request(
            endpoint="me",
            method="GET",
            retry_auth=False
        )

        self.logger.info(
            "Authenticated as %s (%s)",
            me.get("displayName"),
            me.get("userPrincipalName")
        )

        return True

    def is_authenticated(self):

        return (
            self.headers is not None
            and "Authorization" in self.headers
        )

    def ensure_authenticated(self):

        if not self.is_authenticated():
            self.authenticate()

    ###########################################################################
    # Core Request Method
    ###########################################################################

    def request(
        self,
        endpoint,
        method="GET",
        data=None,
        retry_auth=True,
        retry_throttle=True
    ):
        """
        Send request to Graph.
        """

        self.ensure_authenticated()

        url = f"{self.GRAPH_BASE_URL}/{endpoint.lstrip('/')}"

        try:

            response = self.session.request(
                method=method.upper(),
                url=url,
                headers=self.headers,
                json=data,
                timeout=self.timeout
            )

        except requests.RequestException as ex:

            raise GraphApiError(
                f"Network error calling Graph: {ex}"
            ) from ex

        #######################################################################
        # Success
        #######################################################################

        if response.status_code in (200, 201, 202):

            if response.text:
                return response.json()

            return {}

        if response.status_code == 204:
            return {}

        #######################################################################
        # Unauthorized
        #######################################################################

        if response.status_code == 401:

            if retry_auth:

                self.logger.warning(
                    "Received 401. Refreshing token."
                )

                self.authenticate()

                return self.request(
                    endpoint=endpoint,
                    method=method,
                    data=data,
                    retry_auth=False,
                    retry_throttle=retry_throttle
                )

            raise GraphAuthenticationError(
                "Authentication failed after refresh."
            )

        #######################################################################
        # Throttling
        #######################################################################

        if response.status_code == 429:

            if retry_throttle:

                wait_time = int(
                    response.headers.get(
                        "Retry-After",
                        5
                    )
                )

                self.logger.warning(
                    "Graph throttling detected. "
                    "Waiting %s seconds.",
                    wait_time
                )

                time.sleep(wait_time)

                return self.request(
                    endpoint=endpoint,
                    method=method,
                    data=data,
                    retry_auth=retry_auth,
                    retry_throttle=False
                )

        #######################################################################
        # Failure
        #######################################################################

        raise GraphApiError(
            f"Graph API Error\n"
            f"Status: {response.status_code}\n"
            f"Response: {response.text}"
        )

    ###########################################################################
    # Convenience Methods
    ###########################################################################

    def get(self, endpoint):

        return self.request(
            endpoint=endpoint,
            method="GET"
        )

    def post(self, endpoint, data):

        return self.request(
            endpoint=endpoint,
            method="POST",
            data=data
        )

    def patch(self, endpoint, data):

        return self.request(
            endpoint=endpoint,
            method="PATCH",
            data=data
        )

    def delete(self, endpoint):

        return self.request(
            endpoint=endpoint,
            method="DELETE"
        )

    ###########################################################################
    # Pagination Helper
    ###########################################################################

    def get_all(self, endpoint):
        """
        Automatically follows @odata.nextLink and
        returns all records.
        """

        self.ensure_authenticated()

        url = f"{self.GRAPH_BASE_URL}/{endpoint}"

        results = []

        while url:

            response = self.session.get(
                url,
                headers=self.headers,
                timeout=self.timeout
            )

            response.raise_for_status()

            payload = response.json()

            results.extend(
                payload.get("value", [])
            )

            url = payload.get("@odata.nextLink")

        return results