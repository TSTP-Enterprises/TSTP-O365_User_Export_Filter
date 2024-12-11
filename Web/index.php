    <?php
    // Start session to maintain state
    session_start();
    
    // Helper function to check if file is CSV
    function isCSV($file) {
        return $file['type'] === 'text/csv' || 
               strtolower(pathinfo($file['name'], PATHINFO_EXTENSION)) === 'csv';
    }
    
    // Helper function to parse CSV data
    function parseCSV($file) {
        $data = [];
        if (($handle = fopen($file['tmp_name'], "r")) !== FALSE) {
            // Get headers
            $headers = array_map('trim', fgetcsv($handle));
            
            // Read data rows
            while (($row = fgetcsv($handle)) !== FALSE) {
                // Clean up the row data
                $row = array_map('trim', $row);
                // Create associative array with cleaned headers
                $rowData = array_combine($headers, $row);
                // Normalize specific fields
                foreach (['Block credential', 'DirSyncEnabled', 'Password never expires', 'Strong password required'] as $field) {
                    if (isset($rowData[$field])) {
                        $rowData[$field] = strtolower(trim($rowData[$field]));
                    }
                }
                $data[] = $rowData;
            }
            fclose($handle);
        }
        return ['headers' => $headers, 'data' => $data];
    }
    
    // Helper function to apply filters
    function applyFilters($data, $filters) {
        $filtered = [];
    
        // Create a mapping of filter types to their corresponding conditions
        $filterConditions = [
            'Active' => function($row) {
                return isset($row['Block credential']) && $row['Block credential'] === 'false';
            },
            'Sign-In Allowed' => function($row) {
                return isset($row['Block credential']) && $row['Block credential'] === 'false';
            },
            'Sign-In Blocked' => function($row) {
                return isset($row['Block credential']) && $row['Block credential'] === 'true';
            },
            'Licensed' => function($row) {
                return isset($row['Licenses']) && !empty($row['Licenses']);
            },
            'Unlicensed' => function($row) {
                return isset($row['Licenses']) && empty($row['Licenses']);
            },
            'DirSync Enabled' => function($row) {
                return isset($row['DirSyncEnabled']) && $row['DirSyncEnabled'] === 'true';
            },
            'Password Never Expires' => function($row) {
                return isset($row['Password never expires']) && $row['Password never expires'] === 'true';
            },
            'Strong Password Required' => function($row) {
                return isset($row['Strong password required']) && $row['Strong password required'] === 'true';
            },
            'Has City' => function($row) {
                return isset($row['City']) && !empty($row['City']);
            },
            'No City' => function($row) {
                return isset($row['City']) && empty($row['City']);
            },
            'Has Department' => function($row) {
                return isset($row['Department']) && !empty($row['Department']);
            },
            'No Department' => function($row) {
                return isset($row['Department']) && empty($row['Department']);
            },
            // Add more filter conditions as needed
        ];
    
        // Debug logging
        error_log("Applying filters: " . print_r($filters, true));
    
        // Iterate through the data and apply all filters
        foreach ($data as $row) {
            $includeRow = true;
            foreach ($filters as $filter) {
                if (isset($filterConditions[$filter])) {
                    // Debug logging
                    foreach (['Block credential', 'DirSyncEnabled', 'Password never expires', 'Strong password required'] as $field) {
                        if (isset($row[$field])) {
                            error_log("Row {$field} value: '" . $row[$field] . "'");
                        }
                    }
                    if (!$filterConditions[$filter]($row)) {
                        $includeRow = false;
                        break;
                    }
                }
            }
            if ($includeRow) {
                $filtered[] = $row;
            }
        }
    
        // Debug logging
        error_log("Filtered results count: " . count($filtered));
        return $filtered;
    }
    
    // Handle AJAX requests
    if(isset($_FILES['csv_file'])) {
        if(isCSV($_FILES['csv_file'])) {
            $csvData = parseCSV($_FILES['csv_file']);
            $_SESSION['headers'] = $csvData['headers'];
            $_SESSION['data'] = $csvData['data'];
            $_SESSION['selected_columns'] = $csvData['headers'];
            echo json_encode(['success' => true, 'data' => $csvData]);
            exit;
        } else {
            echo json_encode(['success' => false, 'message' => 'Invalid CSV file.']);
            exit;
        }
    }
    
    if(isset($_POST['filters'])) {
        $_SESSION['filters'] = $_POST['filters'];
        $filteredData = applyFilters($_SESSION['data'], $_POST['filters']);
        echo json_encode(['success' => true, 'data' => $filteredData]);
        exit;
    }
    
    if(isset($_POST['columns'])) {
        $_SESSION['selected_columns'] = $_POST['columns'];
        echo json_encode(['success' => true]);
        exit;
    }
    
    ?>
    
    <!DOCTYPE html>
    <html>
    <head>
        <title>Office 365 User Filter</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <!-- Bootstrap CSS with Dark Mode Support -->
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body.light-mode {
                background-color: #f8f9fa;
                color: #212529;
            }
            body.dark-mode {
                background-color: #121212;
                color: #e0e0e0;
            }
            .main-container {
                max-width: 1400px;
                margin: 0 auto;
                padding: 30px;
            }
            .report-header {
                background: linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%);
                color: white;
                padding: 30px;
                margin-bottom: 40px;
                border-radius: 15px;
                box-shadow: 0 6px 8px rgba(0,0,0,0.15);
            }
            .filter-btn.active {
                background-color: #0d6efd;
                color: white;
            }
            .section-card {
                background: white;
                border-radius: 15px;
                padding: 25px;
                margin-bottom: 25px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }
            .column-grid {
                display: grid;
                grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
                gap: 20px;
                margin: 20px 0;
            }
            .column-item {
                padding: 15px;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                display: flex;
                align-items: center;
                background-color: #f1f3f5;
            }
            .table-container {
                max-height: 700px;
                overflow-y: auto;
                margin-top: 25px;
                border-radius: 15px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.05);
                padding: 15px;
                background-color: #ffffff;
            }
            .table th {
                position: sticky;
                top: 0;
                background-color: #f8f9fa;
                z-index: 1;
                border-bottom: 2px solid #dee2e6;
            }
            .export-section {
                display: flex;
                gap: 15px;
                margin: 25px 0;
            }
            .btn-group {
                flex-wrap: wrap;
                gap: 10px;
                margin-top: 15px;
            }
            .dark-mode .section-card {
                background: #1e1e1e;
                border: 1px solid #333;
                box-shadow: 0 4px 6px rgba(255,255,255,0.1);
            }
            .dark-mode .report-header {
                background: linear-gradient(135deg, #343a40 0%, #212529 100%);
            }
            .dark-mode .table th {
                background-color: #343a40;
                color: #e0e0e0;
                border-bottom: 2px solid #495057;
            }
            .dark-mode .btn-outline-primary {
                color: #e0e0e0;
                border-color: #6c757d;
            }
            .dark-mode .btn-outline-primary.active {
                background-color: #0d6efd;
                color: white;
                border-color: #0d6efd;
            }
            .dark-mode .btn-outline-secondary {
                color: #e0e0e0;
                border-color: #6c757d;
            }
            .dark-mode .btn-primary {
                background-color: #0d6efd;
                border-color: #0d6efd;
            }
            .dark-mode .btn-success {
                background-color: #28a745;
                border-color: #28a745;
            }
            .dark-mode .column-item {
                background-color: #2c2c2c;
                color: #e0e0e0;
            }
            @media (max-width: 768px) {
                .column-grid {
                    grid-template-columns: 1fr;
                }
                .btn-group {
                    flex-direction: column;
                }
                .btn-group .btn {
                    width: 100%;
                    margin: 5px 0;
                }
            }
        </style>
    </head>
    <body class="light-mode">
        <div class="main-container">
            <div class="d-flex justify-content-between align-items-center mb-4">
                <div class="report-header text-center flex-grow-1">
                    <h1>Office 365 User Report</h1>
                </div>
                <div class="form-check form-switch ms-4">
                    <input class="form-check-input" type="checkbox" id="darkModeToggle">
                    <label class="form-check-label" for="darkModeToggle">Dark Mode</label>
                </div>
            </div>

            <!-- Tabs Navigation -->
            <ul class="nav nav-tabs mb-4" id="mainTabs" role="tablist">
                <li class="nav-item" role="presentation">
                    <button class="nav-link active" id="filter-tab" data-bs-toggle="tab" data-bs-target="#filter" type="button" role="tab" aria-controls="filter" aria-selected="true">Filter</button>
                </li>
                <li class="nav-item" role="presentation">
                    <button class="nav-link" id="tutorial-tab" data-bs-toggle="tab" data-bs-target="#tutorial" type="button" role="tab" aria-controls="tutorial" aria-selected="false">Tutorial</button>
                </li>
                <li class="nav-item" role="presentation">
                    <button class="nav-link" id="donate-tab" data-bs-toggle="tab" data-bs-target="#donate" type="button" role="tab" aria-controls="donate" aria-selected="false">Donate</button>
                </li>
                <li class="nav-item" role="presentation">
                    <button class="nav-link" id="about-tab" data-bs-toggle="tab" data-bs-target="#about" type="button" role="tab" aria-controls="about" aria-selected="false">About</button>
                </li>
            </ul>
            <div class="tab-content" id="mainTabsContent">
                <!-- Filter Tab -->
                <div class="tab-pane fade show active" id="filter" role="tabpanel" aria-labelledby="filter-tab">
                    <!-- File Upload -->
                    <div class="section-card">
                        <div class="mb-4">
                            <label for="csv_file" class="form-label fs-5">Upload CSV File</label>
                            <input type="file" class="form-control" id="csv_file" accept=".csv">
                        </div>
                    </div>
        
                    <div id="content-section" style="display: none;">
                        <!-- Filters -->
                        <div class="section-card">
                            <h4 class="mb-3">Filters</h4>
                            <div class="btn-group" role="group">
                                <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Active">Active</button>
                                <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Licensed">Licensed</button>
                                <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Unlicensed">Unlicensed</button>
                                <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Sign-In Blocked">Sign-In Blocked</button>
                                <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Sign-In Allowed">Sign-In Allowed</button>
                            </div>
                            <div class="advanced-filters mt-4">
                                <button type="button" class="btn btn-outline-secondary mb-3" id="toggle-advanced-filters">Advanced Filters</button>
                                <div id="advanced-filters-content" style="display: none;">
                                    <div class="btn-group" role="group">
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="DirSync Enabled">DirSync Enabled</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Password Never Expires">Password Never Expires</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Strong Password Required">Strong Password Required</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Has City">Has City</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="No City">No City</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="Has Department">Has Department</button>
                                        <button type="button" class="btn btn-outline-primary filter-btn" data-filter="No Department">No Department</button>
                                        <!-- Additional advanced filters can be added here -->
                                    </div>
                                </div>
                            </div>
                        </div>
        
                        <!-- Column Selection -->
                        <div class="section-card">
                            <h4 class="mb-3">Select Columns</h4>
                            <div class="mb-3">
                                <button class="btn btn-outline-secondary me-2" id="toggle-columns">Show All Columns</button>
                                <button class="btn btn-primary" id="select-default">Select Default Columns</button>
                            </div>
                            <div id="columns-content" style="display: none;">
                                <div class="mb-3">
                                    <button class="btn btn-secondary me-2" id="select-all">Select All</button>
                                    <button class="btn btn-secondary me-2" id="deselect-all">Deselect All</button>
                                </div>
                                <div id="column-grid" class="column-grid">
                                    <!-- Columns will be dynamically added here -->
                                </div>
                            </div>
                        </div>
        
                        <!-- Export Options -->
                        <div class="section-card">
                            <h4 class="mb-3">Export Options</h4>
                            <div class="export-section">
                                <button class="btn btn-success" onclick="exportData('csv')">Export to CSV</button>
                                <button class="btn btn-success" onclick="exportData('txt')">Export to TXT</button>
                                <button class="btn btn-success" onclick="exportData('pdf')">Export to PDF</button>
                            </div>
                        </div>
        
                        <!-- Data Table -->
                        <div class="table-container">
                            <table class="table table-bordered table-hover">
                                <thead>
                                    <tr id="header-row"></tr>
                                </thead>
                                <tbody id="data-body"></tbody>
                            </table>
                        </div>
                    </div>
                </div>
        
                <!-- Tutorial Tab -->
                <div class="tab-pane fade" id="tutorial" role="tabpanel" aria-labelledby="tutorial-tab">
                    <div class="section-card">
                        <h3>How to Use Office 365 User Filter</h3>
                        <ol class="mt-3">
                            <li>Upload your CSV file:
                                <ul>
                                    <li>Click the "Upload CSV File" button</li>
                                    <li>Select your Office 365 user export file</li>
                                    <li>The tool accepts standard O365 admin center exports</li>
                                    <li>File size limit is 10MB for optimal performance</li>
                                </ul>
                            </li>
                            <li>Apply filters to narrow down data:
                                <ul>
                                    <li>Use basic filters like "Active", "Licensed", "Sign-In Blocked"</li>
                                    <li>Apply advanced filters for City, Department, etc.</li>
                                    <li>Combine multiple filters for precise results</li>
                                    <li>Filters work in real-time as you select them</li>
                                    <li>Clear individual filters or reset all at once</li>
                                </ul>
                            </li>
                            <li>Customize your view:
                                <ul>
                                    <li>Show/hide specific columns using "Select Columns"</li>
                                    <li>Use "Select Default Columns" for standard view</li>
                                    <li>Toggle "Show All Columns" to view all fields</li>
                                    <li>Drag columns to reorder them</li>
                                    <li>Sort data by clicking column headers</li>
                                </ul>
                            </li>
                            <li>Export filtered data:
                                <ul>
                                    <li>Choose from CSV, TXT, or PDF formats</li>
                                    <li>Only selected columns and filtered data are exported</li>
                                    <li>PDF exports include formatting and headers</li>
                                    <li>CSV exports are compatible with Excel</li>
                                    <li>TXT exports use tab delimitation</li>
                                </ul>
                            </li>
                            <li>Additional Features:
                                <ul>
                                    <li>Toggle Dark/Light mode for comfortable viewing</li>
                                    <li>Real-time updates as you apply filters</li>
                                    <li>Secure local processing - no data sent to servers</li>
                                    <li>Browser-based calculations for speed</li>
                                    <li>Mobile-responsive design</li>
                                </ul>
                            </li>
                            <li>Best Practices:
                                <ul>
                                    <li>Keep CSV files under 10MB for best performance</li>
                                    <li>Use default columns for standard reporting</li>
                                    <li>Clear filters before applying new ones</li>
                                    <li>Export to PDF for formatted reports</li>
                                    <li>Use CSV for data analysis in Excel</li>
                                </ul>
                            </li>
                        </ol>
                    </div>
                </div>

                <!-- Donate Tab -->
                <div class="tab-pane fade" id="donate" role="tabpanel" aria-labelledby="donate-tab">
                    <div class="section-card text-center">
                        <h3>Support the Development</h3>
                        <p class="mt-3">O365 User Filter is a free and open-source tool. Your support helps maintain and improve it!</p>
                        <div class="mt-4">
                            <a href="https://www.paypal.com/donate/?hosted_button_id=RAAYNUTMHPQQN" target="_blank" class="btn btn-primary btn-lg me-3">
                                <i class="fab fa-paypal me-2"></i>Donate via PayPal
                            </a>
                            <a href="https://github.com/TSTP-Enterprises/TSTP-O365_User_List_Filter" target="_blank" class="btn btn-outline-primary btn-lg">
                                <i class="fab fa-github me-2"></i>Star on GitHub
                            </a>
                        </div>
                        <p class="mt-4">Your contributions help us:</p>
                        <ul class="list-unstyled">
                            <li><i class="fas fa-check text-success me-2"></i>Maintain and update the tool regularly</li>
                            <li><i class="fas fa-check text-success me-2"></i>Add new features and improvements</li>
                            <li><i class="fas fa-check text-success me-2"></i>Provide technical support and documentation</li>
                            <li><i class="fas fa-check text-success me-2"></i>Keep the tool free and open-source</li>
                            <li><i class="fas fa-check text-success me-2"></i>Develop additional admin tools</li>
                        </ul>
                        <div class="mt-4">
                            <h4>Other Ways to Support</h4>
                            <ul class="list-unstyled">
                                <li><i class="fas fa-star text-warning me-2"></i>Star our GitHub repository</li>
                                <li><i class="fas fa-code-branch text-info me-2"></i>Contribute code or report issues</li>
                                <li><i class="fas fa-share-alt text-primary me-2"></i>Share with other administrators</li>
                                <li><i class="fas fa-comment text-success me-2"></i>Provide feedback and suggestions</li>
                            </ul>
                        </div>
                    </div>
                </div>

                <!-- About Tab -->
                <div class="tab-pane fade" id="about" role="tabpanel" aria-labelledby="about-tab">
                    <div class="section-card">
                        <h3>About Office 365 User Filter</h3>
                        <p class="mt-3">O365 User Filter is a powerful web application designed to help administrators efficiently manage and analyze their Office 365 user data. Built with security and ease of use in mind, it processes all data locally in your browser.</p>
                        
                        <h4 class="mt-4">Key Features:</h4>
                        <ul>
                            <li><i class="fas fa-filter me-2"></i>Advanced filtering capabilities with real-time updates</li>
                            <li><i class="fas fa-columns me-2"></i>Customizable column selection and ordering</li>
                            <li><i class="fas fa-file-export me-2"></i>Multiple export formats (CSV, TXT, PDF) with formatting</li>
                            <li><i class="fas fa-shield-alt me-2"></i>Secure local data processing with no server storage</li>
                            <li><i class="fas fa-moon me-2"></i>Dark/Light mode support for extended use</li>
                            <li><i class="fas fa-sync me-2"></i>Real-time filter updates and instant feedback</li>
                            <li><i class="fas fa-mobile-alt me-2"></i>Responsive design for all devices</li>
                            <li><i class="fas fa-chart-bar me-2"></i>Basic data analytics and insights</li>
                        </ul>

                        <h4 class="mt-4">Connect With Us:</h4>
                        <ul>
                            <li><a href="https://tstp.xyz" target="_blank"><i class="fas fa-globe me-2"></i>Official Website - Latest Updates and News</a></li>
                            <li><a href="https://github.com/TSTP-Enterprises/TSTP-O365_User_List_Filter" target="_blank"><i class="fab fa-github me-2"></i>GitHub Repository - Source Code and Issues</a></li>
                            <li><a href="https://www.linkedin.com/company/thesolutions-toproblems" target="_blank"><i class="fab fa-linkedin me-2"></i>LinkedIn - Professional Updates</a></li>
                            <li><a href="https://www.youtube.com/@yourpststudios" target="_blank"><i class="fab fa-youtube me-2"></i>YouTube Channel - Tutorials and Tips</a></li>
                            <li><a href="https://www.facebook.com/profile.php?id=61557162643039" target="_blank"><i class="fab fa-facebook me-2"></i>Facebook - Community Updates</a></li>
                        </ul>

                        <h4 class="mt-4">Contact Us</h4>
                        <p>For support, feedback, or inquiries:</p>
                        <ul>
                            <li><i class="fas fa-envelope me-2"></i>Email: <a href="mailto:support@tstp.xyz">support@tstp.xyz</a></li>
                            <li><i class="fab fa-github me-2"></i>GitHub Issues: <a href="https://github.com/TSTP-Enterprises/TSTP-O365_User_List_Filter/issues" target="_blank">Report a Bug or Request Feature</a></li>
                            <li><i class="fas fa-comments me-2"></i>Discord: <a href="https://discord.gg/tstp" target="_blank">Join Our Community</a></li>
                        </ul>

                        <h4 class="mt-4">Version Information</h4>
                        <ul>
                            <li><i class="fas fa-code me-2"></i>Current Version: 1.0.0</li>
                            <li><i class="fas fa-calendar me-2"></i>Last Updated: January 2024</li>
                            <li><i class="fas fa-file-code me-2"></i>License: MIT</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    
        <!-- Bootstrap JS and Dependencies -->
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
        <script>
            let csvData = null;
            let selectedColumns = [];
            let activeFilters = [];
            const defaultColumns = ['Block credential', 'Display name', 'First name', 'Last name', 'Licenses', 'User principal name'];
    
            // Dark Mode Toggle
            const darkModeToggle = document.getElementById('darkModeToggle');
            darkModeToggle.addEventListener('change', function() {
                if(this.checked) {
                    document.body.classList.remove('light-mode');
                    document.body.classList.add('dark-mode');
                } else {
                    document.body.classList.remove('dark-mode');
                    document.body.classList.add('light-mode');
                }
            });
    
            // File upload handling
            $('#csv_file').change(function(e) {
                const file = e.target.files[0];
                if(!file) return;
    
                const formData = new FormData();
                formData.append('csv_file', file);
    
                $.ajax({
                    url: '',
                    type: 'POST',
                    data: formData,
                    processData: false,
                    contentType: false,
                    success: function(response) {
                        try {
                            const result = JSON.parse(response);
                            if(result.success) {
                                csvData = result.data;
                                selectedColumns = csvData.headers;
                                initializeInterface();
                                $('#content-section').slideDown();
                                // Switch to Filter tab
                                var filterTab = new bootstrap.Tab(document.querySelector('#filter-tab'));
                                filterTab.show();
                            } else {
                                alert(result.message || 'Failed to upload CSV.');
                            }
                        } catch (e) {
                            console.error('Invalid response:', response);
                            alert('An error occurred while processing the file.');
                        }
                    },
                    error: function() {
                        alert('An error occurred during the upload.');
                    }
                });
            });
    
            function initializeInterface() {
                // Initialize column selection
                const columnGrid = $('#column-grid');
                columnGrid.empty();
                
                csvData.headers.forEach(column => {
                    columnGrid.append(`
                        <div class="column-item">
                            <input type="checkbox" class="form-check-input me-2" name="columns[]" 
                                   value="${column}" id="col_${column}" checked>
                            <label class="form-check-label" for="col_${column}">${column}</label>
                        </div>
                    `);
                });
    
                updateTable();
                bindEvents();
            }
    
            function bindEvents() {
                // Column selection events
                $('input[name="columns[]"]').change(function() {
                    selectedColumns = [];
                    $('input[name="columns[]"]:checked').each(function() {
                        selectedColumns.push($(this).val());
                    });
                    updateTable();
                });
    
                // Select/Deselect all
                $('#select-all').click(() => {
                    $('input[name="columns[]"]').prop('checked', true).trigger('change');
                });
    
                $('#deselect-all').click(() => {
                    $('input[name="columns[]"]').prop('checked', false).trigger('change');
                });
    
                // Select default columns
                $('#select-default').click(() => {
                    $('input[name="columns[]"]').each(function() {
                        $(this).prop('checked', defaultColumns.includes($(this).val())).trigger('change');
                    });
                });
    
                // Toggle advanced filters
                $('#toggle-advanced-filters').click(function() {
                    $('#advanced-filters-content').slideToggle();
                });
    
                // Toggle columns visibility
                $('#toggle-columns').click(function() {
                    $('#columns-content').slideToggle();
                    const toggleButton = $(this);
                    toggleButton.text(toggleButton.text() === 'Show All Columns' ? 'Hide Columns' : 'Show All Columns');
                });
    
                // Filter events
                $('.filter-btn').click(function() {
                    $(this).toggleClass('active');
                    activeFilters = [];
                    $('.filter-btn.active').each(function() {
                        activeFilters.push($(this).data('filter'));
                    });
                    updateTable();
                });
            }
    
            function updateTable() {
                if(!csvData) return;
    
                // Update headers
                const headerRow = $('#header-row');
                headerRow.empty();
                selectedColumns.forEach(column => {
                    headerRow.append(`<th>${column}</th>`);
                });
    
                // Update data
                let displayData = csvData.data;
                if(activeFilters.length > 0) {
                    $.post('', {filters: activeFilters}, function(response) {
                        try {
                            const result = JSON.parse(response);
                            if(result.success) {
                                displayData = result.data;
                                updateTableBody(displayData);
                            } else {
                                alert(result.message || 'Failed to apply filters.');
                            }
                        } catch (e) {
                            console.error('Invalid response:', response);
                            alert('An error occurred while applying filters.');
                        }
                    });
                } else {
                    updateTableBody(displayData);
                }
            }
    
            function updateTableBody(displayData) {
                const tbody = $('#data-body');
                tbody.empty();
                
                if(displayData.length === 0){
                    tbody.append('<tr><td colspan="' + selectedColumns.length + '" class="text-center">No data available.</td></tr>');
                    return;
                }
    
                displayData.forEach(row => {
                    const tr = $('<tr>');
                    selectedColumns.forEach(column => {
                        tr.append(`<td>${escapeHtml(row[column] || '')}</td>`);
                    });
                    tbody.append(tr);
                });
            }
    
            // Function to escape HTML to prevent XSS
            function escapeHtml(text) {
                return text.replace(/&/g, "&amp;")
                           .replace(/</g, "&lt;")
                           .replace(/>/g, "&gt;")
                           .replace(/"/g, "&quot;")
                           .replace(/'/g, "&#039;");
            }
    
            // Export functionality
            function exportData(format) {
                if (!csvData) {
                    alert('No data available to export.');
                    return;
                }

                if (format === 'pdf') {
                    const { jsPDF } = window.jspdf;
                    const doc = new jsPDF('l', 'pt', 'a4');
                    const pageWidth = doc.internal.pageSize.width;
                    const margin = 40;
                    const usableWidth = pageWidth - (margin * 2);
                    
                    // Header
                    doc.setFillColor(51, 122, 183);
                    doc.rect(0, 0, pageWidth, 50, 'F');
                    doc.setTextColor(255, 255, 255);
                    doc.setFontSize(22);
                    doc.text('Office 365 User Report', pageWidth / 2, 35, { align: 'center' });
                    
                    doc.setTextColor(0, 0, 0);
                    const data = displayDataToArray(selectedColumns);

                    if (selectedColumns.length <= 7) {
                        // Original table format for 7 or fewer columns
                        doc.autoTable({
                            startY: 70,
                            head: [selectedColumns],
                            body: data,
                            theme: 'grid',
                            styles: {
                                fontSize: 10,
                                cellPadding: 4
                            },
                            headStyles: {
                                fillColor: [240, 240, 240],
                                textColor: [0, 0, 0],
                                fontSize: 12,
                                fontStyle: 'bold'
                            },
                            alternateRowStyles: {
                                fillColor: [245, 245, 245]
                            }
                        });
                    } else {
                        // Card format for 8 or more columns
                        let yPosition = 70;
                        const cellPadding = 5;
                        const cellHeight = 25;
                        
                        // Calculate columns per row (3 columns for card layout)
                        const colsPerRow = 3;
                        const cellWidth = (usableWidth / colsPerRow) - cellPadding;
                        
                        data.forEach((row, rowIndex) => {
                            let currentY = yPosition;
                            const rowStartY = currentY;
                            let maxHeightInRow = 0;
                            
                            // Create cells in grid format
                            for(let i = 0; i < row.length; i++) {
                                const colInRow = i % colsPerRow;
                                const xPos = margin + (colInRow * (cellWidth + cellPadding));
                                
                                if(colInRow === 0 && i !== 0) {
                                    currentY += maxHeightInRow + cellPadding;
                                    maxHeightInRow = 0;
                                }
                                
                                // Draw cell background
                                doc.setFillColor(245, 245, 245);
                                doc.rect(xPos, currentY, cellWidth, cellHeight, 'F');
                                
                                // Draw label and value
                                doc.setFontSize(8);
                                doc.setFont(undefined, 'bold');
                                doc.text(selectedColumns[i], xPos + 5, currentY + 12);
                                doc.setFont(undefined, 'normal');
                                doc.text(row[i], xPos + 5, currentY + 22);
                                
                                maxHeightInRow = Math.max(maxHeightInRow, cellHeight);
                            }
                            
                            // Update yPosition for next user
                            yPosition = currentY + maxHeightInRow + 20;
                            
                            // Add page if needed
                            if(yPosition > doc.internal.pageSize.height - 50) {
                                doc.addPage();
                                yPosition = 40;
                            }
                            
                            // Add separator line
                            doc.setDrawColor(200, 200, 200);
                            doc.line(margin, yPosition - 10, pageWidth - margin, yPosition - 10);
                        });
                    }
                    
                    doc.save('office365_users.pdf');
                } else {
                    // For CSV and TXT, create a Blob and trigger download
                    const data = displayDataToArray(selectedColumns);
                    let content = '';
    
                    if (format === 'csv') {
                        content += selectedColumns.join(",") + "\n";
                        data.forEach(row => {
                            const sanitizedRow = row.map(field => `"${field.replace(/"/g, '""')}"`);
                            content += sanitizedRow.join(",") + "\n";
                        });
                        downloadFile(content, 'office365_users.csv', 'text/csv');
                    } else if (format === 'txt') {
                        content += selectedColumns.join("\t") + "\n";
                        data.forEach(row => {
                            content += row.join("\t") + "\n";
                        });
                        downloadFile(content, 'office365_users.txt', 'text/plain');
                    }
                }
            }

            function displayDataToArray(selectedColumns) {
                const data = [];
                $('#data-body tr').each(function() {
                    const row = [];
                    $(this).find('td').each(function() {
                        row.push($(this).text());
                    });
                    data.push(row);
                });
                return data;
            }

            function downloadFile(content, filename, mimeType) {
                const blob = new Blob([content], { type: mimeType });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        </script>
    </body>
    </html>
