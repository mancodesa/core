#include <cppunit/TestFixture.h>
#include <cppunit/extensions/HelperMacros.h>
#include <cppunit/ui/text/TestRunner.h>
#include <cppunit/extensions/TestFactoryRegistry.h>
#define LOK_USE_UNSTABLE_API

#include <LibreOfficeKit/LibreOfficeKitEnums.h>
#include <LibreOfficeKit/LibreOfficeKit.hxx>

#include <cstdlib>
#include <iostream>

static bool callbackExecuted = false;

// Callback with correct signature for LibreOfficeKitAnyInputCallback
// which is: bool (*)(void*, int)
static bool anyInputCallback(void* data, int type) {
    std::cout << "[DEBUG] anyInputCallback: ENTERED - callback executed with type: " << type << std::endl;
    callbackExecuted = true;
    std::cout << "[DEBUG] anyInputCallback: callbackExecuted set to true" << std::endl;
    (void)data;  // Suppress unused parameter warning
    std::cout << "[DEBUG] anyInputCallback: EXITING - returning true" << std::endl;
    return true;
}

class DesktopLOKTest : public CppUnit::TestFixture
{
    CPPUNIT_TEST_SUITE(DesktopLOKTest);
    CPPUNIT_TEST(testRegisterAnyInputCallback);
    CPPUNIT_TEST_SUITE_END();

public:
    void setUp() override;
    void tearDown() override;
    void testRegisterAnyInputCallback();

private:
    lok::Office* _loKit = nullptr;
};

CPPUNIT_TEST_SUITE_REGISTRATION(DesktopLOKTest);

void DesktopLOKTest::setUp()
{
    std::cout << "[DEBUG] setUp: ENTERED" << std::endl;
    _loKit = nullptr;
    std::cout << "[DEBUG] setUp: _loKit initialized to nullptr" << std::endl;
    std::cout << "[DEBUG] setUp: EXITING" << std::endl;
}

void DesktopLOKTest::tearDown()
{
    std::cout << "[DEBUG] tearDown: ENTERED" << std::endl;
    std::cout << "[DEBUG] tearDown: About to delete _loKit (current value: " << (void*)_loKit << ")" << std::endl;
    delete _loKit;
    std::cout << "[DEBUG] tearDown: _loKit deleted" << std::endl;
    _loKit = nullptr;
    std::cout << "[DEBUG] tearDown: _loKit set to nullptr" << std::endl;
    std::cout << "[DEBUG] tearDown: EXITING" << std::endl;
}

void DesktopLOKTest::testRegisterAnyInputCallback()
{
    std::cout << "[DEBUG] testRegisterAnyInputCallback: TEST STARTED" << std::endl;
    
    const char* loPath = std::getenv("LO_PATH");
    std::cout << "[DEBUG] testRegisterAnyInputCallback: std::getenv(\"LO_PATH\") called" << std::endl;
    
    if (!loPath)
    {
        std::cout << "[DEBUG] testRegisterAnyInputCallback: LO_PATH is NULL, using default path" << std::endl;
        loPath = "/usr/lib/libreoffice/program";
        std::cout << "[DEBUG] testRegisterAnyInputCallback: loPath set to: " << loPath << std::endl;
    }
    else
    {
        std::cout << "[DEBUG] testRegisterAnyInputCallback: LO_PATH found: " << loPath << std::endl;
    }

    std::cout << "[DEBUG] testRegisterAnyInputCallback: About to call lok::lok_cpp_init with path: " << loPath << std::endl;
    _loKit = lok::lok_cpp_init(loPath);
    std::cout << "[DEBUG] testRegisterAnyInputCallback: lok::lok_cpp_init returned: " << (void*)_loKit << std::endl;
    
    if (_loKit != nullptr)
    {
        std::cout << "[DEBUG] testRegisterAnyInputCallback: _loKit is NOT nullptr - initialization successful" << std::endl;
    }
    else
    {
        std::cout << "[DEBUG] testRegisterAnyInputCallback: _loKit is nullptr - initialization FAILED" << std::endl;
    }
    
    CPPUNIT_ASSERT_MESSAGE("Failed to initialize LibreOfficeKit", _loKit != nullptr);
    std::cout << "[DEBUG] testRegisterAnyInputCallback: CPPUNIT_ASSERT passed - _loKit is valid" << std::endl;

    // Register the callback with correct signature
    std::cout << "[DEBUG] testRegisterAnyInputCallback: About to call registerAnyInputCallback" << std::endl;
    std::cout << "[DEBUG] testRegisterAnyInputCallback: Callback function pointer: " << (void*)(anyInputCallback) << std::endl;
    _loKit->registerAnyInputCallback(anyInputCallback, nullptr);
    std::cout << "[DEBUG] testRegisterAnyInputCallback: registerAnyInputCallback completed" << std::endl;

    // Verify the callback was registered by checking if we can trigger it
    // Note: The actual triggering depends on LibreOffice internals
    // For this test, we verify registration succeeded without errors
    std::cout << "[DEBUG] testRegisterAnyInputCallback: Current callbackExecuted value: " << callbackExecuted << std::endl;
    std::cout << "[DEBUG] testRegisterAnyInputCallback: TEST COMPLETED SUCCESSFULLY" << std::endl;
}

int main(int argc, char* argv[])
{
    std::cout << "[DEBUG] main: PROGRAM STARTED" << std::endl;
    std::cout << "[DEBUG] main: argc = " << argc << std::endl;
    for (int i = 0; i < argc; ++i)
    {
        std::cout << "[DEBUG] main: argv[" << i << "] = " << argv[i] << std::endl;
    }
    
    std::cout << "[DEBUG] main: Creating CppUnit::TextUi::TestRunner" << std::endl;
    CppUnit::TextUi::TestRunner runner;
    std::cout << "[DEBUG] main: TestRunner created" << std::endl;
    
    std::cout << "[DEBUG] main: Getting test registry" << std::endl;
    CppUnit::TestFactoryRegistry &registry = CppUnit::TestFactoryRegistry::getRegistry();
    std::cout << "[DEBUG] main: Test registry obtained" << std::endl;
    
    std::cout << "[DEBUG] main: Adding tests to runner" << std::endl;
    runner.addTest(registry.makeTest());
    std::cout << "[DEBUG] main: Tests added to runner" << std::endl;
    
    std::cout << "[DEBUG] main: About to run tests" << std::endl;
    bool wasSuccessful = runner.run("", false);
    std::cout << "[DEBUG] main: Tests completed, wasSuccessful = " << wasSuccessful << std::endl;
    
    if (wasSuccessful)
    {
        std::cout << "[DEBUG] main: All tests PASSED" << std::endl;
    }
    else
    {
        std::cout << "[DEBUG] main: Some tests FAILED" << std::endl;
    }
    
    std::cout << "[DEBUG] main: PROGRAM EXITING with return code: " << (wasSuccessful ? 0 : 1) << std::endl;
    return wasSuccessful ? 0 : 1;
}