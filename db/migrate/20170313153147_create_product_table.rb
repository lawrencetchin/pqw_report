class CreateProductTable < ActiveRecord::Migration[5.0]
  def change
    create_table :products do |t|
      t.string  :product
    end
  end
end
