Rails.application.routes.draw do

  root 'pdf_readers#initiate'
  get 'pdf2excel' => 'pdf_readers#initiate'
  get 'pdf2excel_translate' => 'pdf_readers#translate_initiate'

  post '/transform' => "pdf_readers#transform"
  post '/translate' => "pdf_readers#translate"

  get 'download_pdf_excel' => "pdf_readers#download_excel"
  get '/download_pdf_excel_translate' => 'pdf_readers#download_excel_translate'  

  post '/find_bs_content' => "brightscope#find_bs_content"
  get 'download_bs_excel' => "brightscope#download_excel"
  get 'brightscope' => 'brightscope#initiate'

  # The priority is based upon order of creation: first created -> highest priority.
  # See how all your routes lay out with "rake routes".

  # You can have the root of your site routed with "root"
  # root 'welcome#index'

  # Example of regular route:
  #   get 'products/:id' => 'catalog#view'

  # Example of named route that can be invoked with purchase_url(id: product.id)
  #   get 'products/:id/purchase' => 'catalog#purchase', as: :purchase

  # Example resource route (maps HTTP verbs to controller actions automatically):
  #   resources :products

  # Example resource route with options:
  #   resources :products do
  #     member do
  #       get 'short'
  #       post 'toggle'
  #     end
  #
  #     collection do
  #       get 'sold'
  #     end
  #   end

  # Example resource route with sub-resources:
  #   resources :products do
  #     resources :comments, :sales
  #     resource :seller
  #   end

  # Example resource route with more complex sub-resources:
  #   resources :products do
  #     resources :comments
  #     resources :sales do
  #       get 'recent', on: :collection
  #     end
  #   end

  # Example resource route with concerns:
  #   concern :toggleable do
  #     post 'toggle'
  #   end
  #   resources :posts, concerns: :toggleable
  #   resources :photos, concerns: :toggleable

  # Example resource route within a namespace:
  #   namespace :admin do
  #     # Directs /admin/products/* to Admin::ProductsController
  #     # (app/controllers/admin/products_controller.rb)
  #     resources :products
  #   end
end
