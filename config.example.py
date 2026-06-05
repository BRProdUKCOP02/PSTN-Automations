"""
Gamma API Configuration - EXAMPLE FILE
Copy this file to config.py and fill in your credentials.
config.py is excluded from git via .gitignore.
"""
import os

# =============================================================================
# TEST/UAT ENVIRONMENT CREDENTIALS
# =============================================================================
TEST_USERNAME = os.getenv("GAMMA_TEST_USERNAME", "your_test_username_here")
TEST_PASSWORD = os.getenv("GAMMA_TEST_PASSWORD", "your_test_password_here")

# =============================================================================
# PRODUCTION ENVIRONMENT CREDENTIALS
# =============================================================================
PROD_USERNAME = os.getenv("GAMMA_PROD_USERNAME", "your_prod_username_here")
PROD_PASSWORD = os.getenv("GAMMA_PROD_PASSWORD", "your_prod_password_here")

# =============================================================================
# ENVIRONMENT SELECTION
# =============================================================================
# Set this to choose which environment to use:
#   "TEST" - Uses test/UAT environment (ws-test.gammaoperations.com)
#   "PROD" - Uses production environment (ws.gammaoperations.com)

ENVIRONMENT = "TEST"  # Change to "PROD" for production

# =============================================================================
# DO NOT EDIT BELOW THIS LINE
# =============================================================================

def get_credentials():
    """Get the credentials for the selected environment"""
    if ENVIRONMENT.upper() in ("TEST", "UAT"):
        return {
            'username': TEST_USERNAME,
            'password': TEST_PASSWORD,
            'use_production': False,
            'environment_name': 'TEST/UAT'
        }
    elif ENVIRONMENT.upper() == "PROD":
        return {
            'username': PROD_USERNAME,
            'password': PROD_PASSWORD,
            'use_production': True,
            'environment_name': 'PRODUCTION'
        }
    else:
        raise ValueError(f"Invalid ENVIRONMENT: {ENVIRONMENT}. Must be 'TEST' or 'PROD'")

def get_checker():
    """Get an initialized GammaAccessSuitabilityChecker with the selected credentials"""
    from bb_suitability_checker import GammaAccessSuitabilityChecker

    creds = get_credentials()

    return GammaAccessSuitabilityChecker(
        username=creds['username'],
        password=creds['password'],
        use_production=creds['use_production'],
        auto_refresh=True
    )
