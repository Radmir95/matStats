class User < ActiveRecord::Base

  before_create :create_remember_token
  before_save { self.email = email.downcase }
  validates :first_name, presence: true, :presence => {:message => 'Вы не ввели имя.'}
  validates :last_name, presence: true, :presence => {:message => 'Вы не ввели фамилию.'}
  validates :email, email_format: { message: "Неверная почта." },  uniqueness: { case_sensitive: false }

  has_secure_password
  validates :password, length: { minimum: 6 }
  validates :birth_date, presence: true, :presence => {:message => 'Вы не ввели дату рождения.(Пример: 01.01.2000)'}

  def User.new_remember_token
    SecureRandom.urlsafe_base64
  end

  def User.encrypt(token)
    Digest::SHA1.hexdigest(token.to_s)
  end

  private

  def create_remember_token
    self.remember_token = User.encrypt(User.new_remember_token)
  end

end
