class CreateRategroups < ActiveRecord::Migration[5.0]
  def change
    create_table :rategroups do |t|
      t.string  :rategroup
    end
  end
end
