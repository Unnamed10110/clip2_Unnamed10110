cmake_minimum_required(VERSION 3.15)
project(clip2)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Enable Unicode
add_definitions(-DUNICODE -D_UNICODE)

# Windows-specific settings
if(WIN32)
    set(CMAKE_RC_COMPILER_INIT windres)
    enable_language(RC)
    
    # Add resource file
    set(RESOURCE_FILE ClipboardManager.rc)
    
    add_executable(${PROJECT_NAME} WIN32
        ClipboardManager.cpp
        ClipboardManager.h
        ${RESOURCE_FILE}
        ico2.ico
        click.mp3
    )
    
    # Copy click.mp3 to build directory for fallback
    add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
        "${CMAKE_SOURCE_DIR}/click.mp3"
        "$<TARGET_FILE_DIR:${PROJECT_NAME}>/click.mp3"
    )
    
    # Link necessary libraries
    set(CLIP2_LIBS user32 gdi32 shell32 winmm comctl32 shlwapi ole32 oleaut32)
    # UIAutomation: "Copy from focused control" when app never puts content on clipboard
    set(UIAUTOMATION_SEARCH_PATHS "")
    if(DEFINED ENV{WindowsSdkDir})
        file(GLOB _winsdk_lib_vers "$ENV{WindowsSdkDir}Lib/*")
        foreach(_ver ${_winsdk_lib_vers})
            list(APPEND UIAUTOMATION_SEARCH_PATHS "${_ver}/um/x64" "${_ver}/um/x86")
        endforeach()
    endif()
    file(GLOB _kits_lib_vers "C:/Program Files (x86)/Windows Kits/10/Lib/*")
    foreach(_ver ${_kits_lib_vers})
        if(IS_DIRECTORY "${_ver}")
            list(APPEND UIAUTOMATION_SEARCH_PATHS "${_ver}/um/x64" "${_ver}/um/x86")
        endif()
    endforeach()
    # Explicit fallbacks for common SDK versions (Windows SDK provides UIAutomationCore.lib)
    list(APPEND UIAUTOMATION_SEARCH_PATHS
        "C:/Program Files (x86)/Windows Kits/10/Lib/10.0.26100.0/um/x64"
        "C:/Program Files (x86)/Windows Kits/10/Lib/10.0.26100.0/um/x86"
        "C:/Program Files (x86)/Windows Kits/10/Lib/10.0.22621.0/um/x64"
        "C:/Program Files (x86)/Windows Kits/10/Lib/10.0.22621.0/um/x86"
    )
    find_library(UIAUTOMATION_LIB NAMES UIAutomation UIAutomationCore PATHS ${UIAUTOMATION_SEARCH_PATHS} NO_DEFAULT_PATH)
    if(UIAUTOMATION_LIB)
        list(APPEND CLIP2_LIBS ${UIAUTOMATION_LIB})
        target_compile_definitions(${PROJECT_NAME} PRIVATE HAVE_UIAUTOMATION=1)
        message(STATUS "UIAutomation found: ${UIAUTOMATION_LIB}")
    else()
        message(STATUS "UIAutomation not found; 'Copy from focused control' will be disabled")
    endif()
    target_link_libraries(${PROJECT_NAME} ${CLIP2_LIBS})
    
    # Set output directory
    set_target_properties(${PROJECT_NAME} PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}
    )
endif()

