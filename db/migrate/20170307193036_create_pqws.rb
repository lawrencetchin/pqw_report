class CreatePqws < ActiveRecord::Migration[5.0]
  def change
    create_table :pqws do |t|
        t.string  :pqw_name
    end
  end
end
