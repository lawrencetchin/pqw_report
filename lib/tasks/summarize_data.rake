  
namespace :summarize_data do
    
  desc "Run PQW report"
  task :excel_report => :environment do
    
    puts "Creating report..."
    PqwExcel.update_analysis
    puts "Report created."
  end
    
end