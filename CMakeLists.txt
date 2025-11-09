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
        misc02.ico
    )
    
    # Link necessary libraries
    target_link_libraries(${PROJECT_NAME} 
        user32 
        gdi32 
        shell32 
        winmm
        comctl32
        shlwapi
        ole32
    )
    
    # Set output directory
    set_target_properties(${PROJECT_NAME} PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}
    )
endif()

