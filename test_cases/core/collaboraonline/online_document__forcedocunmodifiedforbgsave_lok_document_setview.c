#include <cppunit/TestFixture.h>
#include <cppunit/extensions/HelperMacros.h>
#include <cppunit/ui/text/TestRunner.h>
#include <cppunit/extensions/TestFactoryRegistry.h>
#define LOK_USE_UNSTABLE_API

#include <LibreOfficeKit/LibreOfficeKitEnums.h>
#include <LibreOfficeKit/LibreOfficeKit.hxx>

#include <cstdlib>
#include <iostream>
#include <fstream>
#include <string>
#include <unistd.h>

// Sequence: lok::Document::setView,lok::Document::registerCallback,lok::Office::registerCallback,lok::Document::destroyView,lok::Document::getViewsCount

static void documentCallback(int /*type*/, const char* /*payload*/, void* /*data*/) {
    // Document callback handler
    std::cout << "[DEBUG] documentCallback invoked" << std::endl;
}

static void officeCallback(int /*type*/, const char* /*payload*/, void* /*data*/) {
    // Office callback handler
    std::cout << "[DEBUG] officeCallback invoked" << std::endl;
}

class DesktopLOKTest : public CppUnit::TestFixture
{
    CPPUNIT_TEST_SUITE(DesktopLOKTest);
    CPPUNIT_TEST(testSetView);
    CPPUNIT_TEST_SUITE_END();

public:
    void setUp() override;
    void tearDown() override;
    void testSetView();

private:
    lok::Office* _loKit = nullptr;
    lok::Document* _loKitDocument = nullptr;
    std::string createTestDocument();
};

CPPUNIT_TEST_SUITE_REGISTRATION(DesktopLOKTest);

void DesktopLOKTest::setUp()
{
    std::cout << "[DEBUG] setUp() - START" << std::endl;
    _loKit = nullptr;
    std::cout << "[DEBUG] setUp() - _loKit set to nullptr" << std::endl;
    _loKitDocument = nullptr;
    std::cout << "[DEBUG] setUp() - _loKitDocument set to nullptr" << std::endl;
    std::cout << "[DEBUG] setUp() - END" << std::endl;
}

void DesktopLOKTest::tearDown()
{
    std::cout << "[DEBUG] tearDown() - START" << std::endl;
    if (_loKitDocument)
    {
        std::cout << "[DEBUG] tearDown() - _loKitDocument is not null, deleting..." << std::endl;
        delete _loKitDocument;
        _loKitDocument = nullptr;
        std::cout << "[DEBUG] tearDown() - _loKitDocument deleted and set to nullptr" << std::endl;
    }
    else
    {
        std::cout << "[DEBUG] tearDown() - _loKitDocument is null, skipping delete" << std::endl;
    }
    std::cout << "[DEBUG] tearDown() - deleting _loKit..." << std::endl;
    delete _loKit;
    _loKit = nullptr;
    std::cout << "[DEBUG] tearDown() - _loKit deleted and set to nullptr" << std::endl;
    std::cout << "[DEBUG] tearDown() - END" << std::endl;
}

std::string DesktopLOKTest::createTestDocument()
{
    std::cout << "[DEBUG] createTestDocument() - START" << std::endl;
    
    // Get current working directory
    char cwd[1024];
    if (getcwd(cwd, sizeof(cwd)) != nullptr) {
        std::cout << "[DEBUG] createTestDocument() - Current working directory: " << cwd << std::endl;
    }
    
    // Create a minimal ODT file (ZIP-based format)
    std::string testDocPath = "/tmp/test_document.odt";
    std::cout << "[DEBUG] createTestDocument() - Creating test document at: " << testDocPath << std::endl;
    
    // Create a minimal valid ODT structure
    // For simplicity, we'll create a minimal XML-based document
    // ODT files are ZIP archives, but we'll create a simple flat XML format
    
    // Create a flat XML ODT document
    std::string flatOdtPath = "/tmp/test_document.fodt";
    std::ofstream outFile(flatOdtPath);
    
    if (!outFile.is_open()) {
        std::cout << "[DEBUG] createTestDocument() - Failed to create file: " << flatOdtPath << std::endl;
        return "";
    }
    
    // Write minimal flat ODT XML content
    outFile << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    outFile << "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ";
    outFile << "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" ";
    outFile << "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" ";
    outFile << "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" ";
    outFile << "office:mimetype=\"application/vnd.oasis.opendocument.text\" ";
    outFile << "office:version=\"1.2\">\n";
    outFile << "  <office:automatic-styles/>\n";
    outFile << "  <office:body>\n";
    outFile << "    <office:text>\n";
    outFile << "      <text:p>Test Document</text:p>\n";
    outFile << "    </office:text>\n";
    outFile << "  </office:body>\n";
    outFile << "</office:document>\n";
    
    outFile.close();
    
    std::cout << "[DEBUG] createTestDocument() - Test document created successfully at: " << flatOdtPath << std::endl;
    std::cout << "[DEBUG] createTestDocument() - END" << std::endl;
    
    return flatOdtPath;
}

void DesktopLOKTest::testSetView()
{
    std::cout << "[DEBUG] testSetView() - START" << std::endl;
    
    const char* loPath = std::getenv("LO_PATH");
    std::cout << "[DEBUG] testSetView() - Retrieved LO_PATH from environment" << std::endl;
    if (!loPath)
    {
        loPath = "/usr/lib/libreoffice/program";
        std::cout << "[DEBUG] testSetView() - LO_PATH not set, using default: " << loPath << std::endl;
    }
    else
    {
        std::cout << "[DEBUG] testSetView() - LO_PATH set to: " << loPath << std::endl;
    }

    std::cout << "[DEBUG] testSetView() - Calling lok::lok_cpp_init()..." << std::endl;
    _loKit = lok::lok_cpp_init(loPath);
    std::cout << "[DEBUG] testSetView() - lok::lok_cpp_init() returned: " << (void*)_loKit << std::endl;
    CPPUNIT_ASSERT_MESSAGE("Failed to initialize LibreOfficeKit", _loKit != nullptr);
    std::cout << "[DEBUG] testSetView() - LibreOfficeKit initialized successfully" << std::endl;

    // Register office callback
    std::cout << "[DEBUG] testSetView() - Calling _loKit->registerCallback()..." << std::endl;
    _loKit->registerCallback(officeCallback, nullptr);
    std::cout << "[DEBUG] testSetView() - Office callback registered successfully" << std::endl;

    // Create test document
    std::cout << "[DEBUG] testSetView() - Creating test document..." << std::endl;
    std::string testDocPath = createTestDocument();
    
    if (testDocPath.empty()) {
        std::cout << "[DEBUG] testSetView() - Failed to create test document" << std::endl;
        CPPUNIT_FAIL("Failed to create test document");
    }
    
    std::cout << "[DEBUG] testSetView() - Test document path: " << testDocPath << std::endl;

    std::cout << "[DEBUG] testSetView() - Calling _loKit->documentLoad()..." << std::endl;
    _loKitDocument = _loKit->documentLoad(testDocPath.c_str());
    std::cout << "[DEBUG] testSetView() - _loKit->documentLoad() returned: " << (void*)_loKitDocument << std::endl;
    CPPUNIT_ASSERT_MESSAGE("Failed to load document via LibreOfficeKit", _loKitDocument != nullptr);
    std::cout << "[DEBUG] testSetView() - Document loaded successfully" << std::endl;

    // Register document callback
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->registerCallback()..." << std::endl;
    _loKitDocument->registerCallback(documentCallback, nullptr);
    std::cout << "[DEBUG] testSetView() - Document callback registered successfully" << std::endl;

    // Get initial view count
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->getViewsCount()..." << std::endl;
    int initialViewCount = _loKitDocument->getViewsCount();
    std::cout << "[DEBUG] testSetView() - Initial view count: " << initialViewCount << std::endl;
    CPPUNIT_ASSERT_MESSAGE("Initial view count should be at least 1", initialViewCount >= 1);
    std::cout << "[DEBUG] testSetView() - Initial view count assertion passed" << std::endl;

    // Create an additional view
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->createView()..." << std::endl;
    int newViewId = _loKitDocument->createView();
    std::cout << "[DEBUG] testSetView() - New view created with ID: " << newViewId << std::endl;
    CPPUNIT_ASSERT_MESSAGE("Failed to create new view", newViewId >= 0);
    std::cout << "[DEBUG] testSetView() - New view creation assertion passed" << std::endl;

    // Set view to view ID 0
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->setView(0)..." << std::endl;
    _loKitDocument->setView(0);
    std::cout << "[DEBUG] testSetView() - setView(0) called" << std::endl;

    // Verify the view was set correctly
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->getView()..." << std::endl;
    int viewAfterSet = _loKitDocument->getView();
    std::cout << "[DEBUG] testSetView() - Current view after setView(0): " << viewAfterSet << std::endl;
    CPPUNIT_ASSERT_EQUAL_MESSAGE("setView(0) failed", 0, viewAfterSet);
    std::cout << "[DEBUG] testSetView() - setView(0) assertion passed" << std::endl;

    // Set view to the newly created view
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->setView(" << newViewId << ")..." << std::endl;
    _loKitDocument->setView(newViewId);
    std::cout << "[DEBUG] testSetView() - setView(" << newViewId << ") called" << std::endl;

    // Verify the view was set correctly
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->getView()..." << std::endl;
    int viewAfterSet2 = _loKitDocument->getView();
    std::cout << "[DEBUG] testSetView() - Current view after setView(" << newViewId << "): " << viewAfterSet2 << std::endl;
    CPPUNIT_ASSERT_EQUAL_MESSAGE("setView(newViewId) failed", newViewId, viewAfterSet2);
    std::cout << "[DEBUG] testSetView() - setView(newViewId) assertion passed" << std::endl;

    // Get view count after creating view
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->getViewsCount()..." << std::endl;
    int viewCountAfterCreate = _loKitDocument->getViewsCount();
    std::cout << "[DEBUG] testSetView() - View count after create: " << viewCountAfterCreate << " (expected: " << (initialViewCount + 1) << ")" << std::endl;
    CPPUNIT_ASSERT_EQUAL_MESSAGE("View count should increase by 1", initialViewCount + 1, viewCountAfterCreate);
    std::cout << "[DEBUG] testSetView() - View count after create assertion passed" << std::endl;

    // Destroy the newly created view
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->destroyView(" << newViewId << ")..." << std::endl;
    _loKitDocument->destroyView(newViewId);
    std::cout << "[DEBUG] testSetView() - destroyView(" << newViewId << ") called" << std::endl;

    // Get final view count
    std::cout << "[DEBUG] testSetView() - Calling _loKitDocument->getViewsCount()..." << std::endl;
    int finalViewCount = _loKitDocument->getViewsCount();
    std::cout << "[DEBUG] testSetView() - Final view count: " << finalViewCount << " (expected: " << initialViewCount << ")" << std::endl;
    CPPUNIT_ASSERT_EQUAL_MESSAGE("View count should return to initial", initialViewCount, finalViewCount);
    std::cout << "[DEBUG] testSetView() - Final view count assertion passed" << std::endl;
    
    std::cout << "[DEBUG] testSetView() - END (SUCCESS)" << std::endl;
}

int main(int /*argc*/, char* /*argv*/[])
{
    std::cout << "[DEBUG] main() - START" << std::endl;
    
    CppUnit::TextUi::TestRunner runner;
    std::cout << "[DEBUG] main() - TestRunner created" << std::endl;
    
    CppUnit::TestFactoryRegistry &registry = CppUnit::TestFactoryRegistry::getRegistry();
    std::cout << "[DEBUG] main() - TestFactoryRegistry obtained" << std::endl;
    
    runner.addTest(registry.makeTest());
    std::cout << "[DEBUG] main() - Tests added to runner" << std::endl;
    
    std::cout << "[DEBUG] main() - Running tests..." << std::endl;
    bool wasSuccessful = runner.run("", false);
    std::cout << "[DEBUG] main() - Tests completed. Result: " << (wasSuccessful ? "SUCCESS" : "FAILURE") << std::endl;
    
    std::cout << "[DEBUG] main() - END (returning " << (wasSuccessful ? 0 : 1) << ")" << std::endl;
    return wasSuccessful ? 0 : 1;
}
