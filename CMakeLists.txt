cmake_minimum_required(VERSION 3.20)

set(CMAKE_INSTALL_PREFIX "${CMAKE_CURRENT_BINARY_DIR}/install" CACHE PATH "Install path prefix, default is <build>/install")

project(MyCOMTest VERSION 1.0.0.0)

set(_CMAKE_TOOLCHAIN_PREFIX "${_CMAKE_TOOLCHAIN_PREFIX}" CACHE STRING "Toolchain prefix")

list(APPEND _CMAKE_IDL_COMPILER_LIST "widl;midl")

list(APPEND _CMAKE_IDL_COMPILER_LIST "${_CMAKE_TOOLCHAIN_PREFIX}widl")

find_program(CMAKE_IDL_COMPILER NAMES ${_CMAKE_IDL_COMPILER_LIST} DOC "IDL compiler")

message(STATUS "Using IDL compiler: ${CMAKE_IDL_COMPILER}")

if(NOT CMAKE_IDL_COMPILER)
    message(FATAL_ERROR "Could not find an IDL compiler. Please install 'widl' or 'midl'.")
endif()

cmake_path(GET CMAKE_IDL_COMPILER PARENT_PATH CMAKE_IDL_COMPILER_DIR)

message(STATUS "IDL compiler directory: ${CMAKE_IDL_COMPILER_DIR}")

add_library(debugOptions INTERFACE)

target_link_options(debugOptions INTERFACE
    $<$<CXX_COMPILER_ID:Clang>:-Wl,--pdb=>
)

if(CMAKE_CXX_COMPILER_VERSION VERSION_GREATER_EQUAL 14.0.0)
    target_link_options(debugOptions INTERFACE
        $<$<CXX_COMPILER_ID:GNU>:-Wl,--pdb=>
    )
    target_compile_options(debugOptions INTERFACE
        $<$<CXX_COMPILER_ID:GNU>:-gcodeview>
    )
endif()

target_compile_options(debugOptions INTERFACE
    $<$<CXX_COMPILER_ID:Clang>:-gcodeview>
    $<$<CXX_COMPILER_ID:MSVC>:/Zi>
)

add_library(rtlibOptions INTERFACE)

if(MSVC)
    target_compile_options(rtlibOptions INTERFACE
        $<$<CONFIG:Debug>:/MTd>      # Debug 版本使用 /MTd
        $<$<CONFIG:Release>:/MT>     # Release 版本使用 /MT
        $<$<CONFIG:RelWithDebInfo>:/MT>
        $<$<CONFIG:MinSizeRel>:/MT>
    )
else()
    target_link_options(rtlibOptions INTERFACE
        $<$<CXX_COMPILER_ID:GNU>:-static>
        $<$<CXX_COMPILER_ID:Clang>:-static>
    )
endif()

set(IDL_FILE_NAME "SimpleCOM")

set(IDL_FILE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/${IDL_FILE_NAME}.idl")

set(GENERATED_DIR "${CMAKE_CURRENT_BINARY_DIR}/Generated")

set(IDL_OUTPUT_HDR "${GENERATED_DIR}/${IDL_FILE_NAME}_i.h")
set(IDL_OUTPUT_GUID "${GENERATED_DIR}/${IDL_FILE_NAME}_i.c")
set(IDL_OUTPUT_TLB "${GENERATED_DIR}/${IDL_FILE_NAME}.tlb")
set(IDL_OUTPUT_PROXY "${GENERATED_DIR}/${IDL_FILE_NAME}_p.c")
set(IDL_OUTPUT_DLLDATA "${GENERATED_DIR}/dlldata.c")

set(IDL_OUTPUT_FILES
    "${IDL_OUTPUT_HDR}"
    "${IDL_OUTPUT_GUID}"
    "${IDL_OUTPUT_TLB}"
    "${IDL_OUTPUT_PROXY}"
    "${IDL_OUTPUT_DLLDATA}"
)

file(MAKE_DIRECTORY "${GENERATED_DIR}")

if(MINGW)
    # widl set libdir use  (strstr(arg[0],'/')!!! not '\\') append ../include,if windows system arg[0] not contain '/' ,it will use ../include as libdir,
    # so workdirectory may be important,it is a bug， or set -L
    add_custom_command(
        OUTPUT "${IDL_OUTPUT_HDR}" "${IDL_OUTPUT_GUID}" "${IDL_OUTPUT_TLB}" "${IDL_OUTPUT_PROXY}"
        COMMAND "${CMAKE_IDL_COMPILER}" -h -H "${IDL_OUTPUT_HDR}" 
            -u -U "${IDL_OUTPUT_GUID}"
            -t -T "${IDL_OUTPUT_TLB}"
            -p -P "${IDL_OUTPUT_PROXY}"
            "${IDL_FILE_PATH}"
        WORKING_DIRECTORY "${CMAKE_IDL_COMPILER_DIR}"
        DEPENDS "${IDL_FILE_PATH}"
        COMMENT "Compiling IDL file with proxy/stub: ${IDL_FILE_PATH}"
    )
    add_custom_command(
        OUTPUT "${IDL_OUTPUT_DLLDATA}"
        COMMAND "${CMAKE_IDL_COMPILER}" --dlldata-only --dlldata "${IDL_OUTPUT_DLLDATA}"
            "${IDL_FILE_PATH}"
        WORKING_DIRECTORY "${CMAKE_IDL_COMPILER_DIR}"
        DEPENDS "${IDL_FILE_PATH}"
        COMMENT "Compiling IDL file with dlldata: ${IDL_FILE_PATH}"
    )
else()
    add_custom_command(
        OUTPUT ${IDL_OUTPUT_FILES}
        COMMAND ${CMAKE_IDL_COMPILER} /h ${IDL_FILE_NAME}_i.h /iid ${IDL_FILE_NAME}_i.c /proxy ${IDL_FILE_NAME}_p.c /dlldata dlldata.c /tlb ${IDL_FILE_NAME}.tlb ${IDL_FILE_PATH}
        WORKING_DIRECTORY ${GENERATED_DIR}
        DEPENDS ${IDL_FILE_PATH}
        VERBATIM
    )
endif()

if(NOT PROJECT_VERSION_TWEAK)
    set(PROJECT_VERSION_TWEAK 0)
endif()

set(SIMPLECOM_VERSION_MAJOR ${PROJECT_VERSION_MAJOR})
set(SIMPLECOM_VERSION_MINOR ${PROJECT_VERSION_MINOR})
set(SIMPLECOM_VERSION_PATCH ${PROJECT_VERSION_PATCH})
set(SIMPLECOM_VERSION_TWEAK ${PROJECT_VERSION_TWEAK})

string(TIMESTAMP BUILD_DATE "%Y-%m-%d %H:%M:%S" UTC)
string(TIMESTAMP COPYRIGHT_YEAR "%Y")

configure_file(SimpleCOM.rc.in ${CMAKE_CURRENT_BINARY_DIR}/SimpleCOM.rc @ONLY)

option(_MERGE_PROXYSTUB "Merge proxy/stub into main DLL" ON)

add_library(SimpleCOM SHARED SimpleCOM.cpp dllmain.cpp SimpleCOM.def
    ${IDL_OUTPUT_HDR}
    ${IDL_OUTPUT_GUID}
    ${CMAKE_CURRENT_BINARY_DIR}/SimpleCOM.rc
    $<$<BOOL:${_MERGE_PROXYSTUB}>:${IDL_OUTPUT_DLLDATA} ${IDL_OUTPUT_PROXY}>
)

set_target_properties(SimpleCOM PROPERTIES 
    PREFIX ""
    IMPORT_PREFIX ""
    IMPORT_SUFFIX ".lib")

target_compile_definitions(SimpleCOM PRIVATE
    _ATL_APARTMENT_THREADED
    _WINDLL
    _USRDLL
    $<$<BOOL:${_MERGE_PROXYSTUB}>:_MERGE_PROXYSTUB>
    $<$<BOOL:${_MERGE_PROXYSTUB}>:REGISTER_PROXY_DLL>
    $<$<BOOL:${_MERGE_PROXYSTUB}>:$<IF:$<CXX_COMPILER_ID:MSVC>,,PROXY_DELEGATION>>
    $<$<BOOL:${_MERGE_PROXYSTUB}>:ENTRY_PREFIX=Prx>
)

target_compile_options(SimpleCOM PRIVATE
    $<$<CXX_COMPILER_ID:MSVC>:/utf-8>
#clang define and declare need declare two times, but gcc/msvc just declare once, so we use this to ignore the warning
    $<$<CXX_COMPILER_ID:Clang>:-Wno-missing-exception-spec>
    $<$<AND:$<CXX_COMPILER_ID:Clang>,$<COMPILE_LANGUAGE:CXX>>:-Wno-invalid-offsetof>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<COMPILE_LANGUAGE:CXX>>:-Wno-invalid-offsetof>
)

target_include_directories(SimpleCOM PRIVATE "${GENERATED_DIR}"
)

target_link_libraries(SimpleCOM PRIVATE $<$<CONFIG:Debug>:debugOptions> rtlibOptions)

target_link_libraries(SimpleCOM PRIVATE
    ole32
    oleaut32
    uuid
    advapi32
    shlwapi
    rpcrt4
)

add_executable(TestClient TestClient.cpp
    ${IDL_OUTPUT_HDR}
    ${IDL_OUTPUT_GUID}
)

target_compile_options(TestClient PRIVATE
    $<$<CXX_COMPILER_ID:MSVC>:/utf-8>
)

target_link_libraries(TestClient PRIVATE $<$<CONFIG:Debug>:debugOptions> rtlibOptions)

target_include_directories(TestClient PRIVATE "${GENERATED_DIR}")

# 创建代理存根 DLL
add_library(SimpleComPS SHARED 
    ${IDL_OUTPUT_PROXY}
    ${IDL_OUTPUT_DLLDATA}
    ${IDL_OUTPUT_GUID}
    SimpleComPS.def
)

set_target_properties(SimpleComPS PROPERTIES 
    PREFIX ""
    IMPORT_PREFIX ""
    IMPORT_SUFFIX ".lib"
)

target_compile_definitions(SimpleComPS PRIVATE
    WIN32
    _WINDLL
    _USRDLL
    REGISTER_PROXY_DLL
    $<IF:$<CXX_COMPILER_ID:MSVC>,,PROXY_DELEGATION>
)

target_compile_options(SimpleComPS PRIVATE
    $<$<CXX_COMPILER_ID:MSVC>:/utf-8>
)

target_include_directories(SimpleComPS PRIVATE 
    "${GENERATED_DIR}"
)

target_link_libraries(SimpleComPS PRIVATE
    rpcrt4
    oleaut32
    uuid
    ole32
)

target_link_libraries(SimpleComPS PRIVATE $<$<CONFIG:Debug>:debugOptions> rtlibOptions)

add_executable(TestAggregation TestAggregation.cpp
)

target_compile_options(TestAggregation PRIVATE
    $<$<CXX_COMPILER_ID:MSVC>:/utf-8>
    $<$<AND:$<CXX_COMPILER_ID:Clang>,$<COMPILE_LANGUAGE:CXX>>:-Wno-invalid-offsetof>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<COMPILE_LANGUAGE:CXX>>:-Wno-invalid-offsetof>
)

target_compile_definitions(TestAggregation PRIVATE _ATL_APARTMENT_THREADED)

target_link_libraries(TestAggregation PRIVATE $<$<CONFIG:Debug>:debugOptions> rtlibOptions)

#install project
if (_MERGE_PROXYSTUB)
    set(REGISTER_PROXYSTUB 0)
else()
    set(REGISTER_PROXYSTUB 1)
endif()
configure_file(register.bat.in ${CMAKE_CURRENT_BINARY_DIR}/register.bat @ONLY  NEWLINE_STYLE WIN32)
configure_file(unregister.bat.in ${CMAKE_CURRENT_BINARY_DIR}/unregister.bat @ONLY  NEWLINE_STYLE WIN32)

install(TARGETS SimpleCOM
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

if (NOT _MERGE_PROXYSTUB)
    install(TARGETS SimpleComPS
        RUNTIME DESTINATION bin
        LIBRARY DESTINATION lib
        ARCHIVE DESTINATION lib
    )
endif()

install(FILES ${IDL_OUTPUT_TLB} ${IDL_FILE_PATH} ${IDL_OUTPUT_HDR} ${IDL_OUTPUT_GUID} DESTINATION include)

install(FILES SimpleCOM.vbs $<TARGET_FILE:TestAggregation> $<TARGET_FILE:TestClient> DESTINATION test)

install(FILES ${CMAKE_CURRENT_BINARY_DIR}/register.bat ${CMAKE_CURRENT_BINARY_DIR}/unregister.bat DESTINATION bin)

set(CPACK_GENERATOR "ZIP")

set(CPACK_PACKAGE_FILE_NAME "${PROJECT_NAME}-${PROJECT_VERSION}-${CMAKE_SYSTEM_NAME}")

include(CPack)