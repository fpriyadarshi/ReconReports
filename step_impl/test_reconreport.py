from getgauge.python import step, Messages, after_suite, after_spec, before_suite, after_step, before_step
from step_impl import reconreport


@step("Run the Recon Report")
def run_recon_report():
    rr = reconreport.Test_URL()
    rr.test_open_url()
    rr.test_loginSalesForce()
    rr.test_ExportReports()
    rr.test_sendMail()
